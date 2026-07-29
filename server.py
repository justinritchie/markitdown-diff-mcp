#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "fastmcp>=3.2,<4",
#   "markitdown[all]>=0.0.1a4,<1",
# ]
# ///
#
# Ceilings on everything, and run LOCALLY via `uv run --script`, never via
# `uvx --from git+…`. That pattern installs the repo as a DEPENDENCY, so uv
# resolves fresh from the loosest specs and ignores any committed lockfile —
# which is how the cloudflare connector died on 2026-07-28 when mcp 2.0.0
# shipped. Local script + pinned deps means an upstream release can't break a
# working connector between one launch and the next. See tasks/lessons.md.
"""
markitdown-diff — local MCP for comparing Markdown, text, and converted documents.

Architecture (single file, all logic clearly sectioned):
  MCP client
    -> markitdown-diff server
      -> input resolver
      -> converter dispatcher (markitdown | docling | pandoc)
      -> markdown/text normalizer
      -> diff engine (difflib)
      -> output renderer (json | unified | markdown)

Converters:
  - markitdown (default, in-process Python lib; great for DOCX, weak on complex PDF)
  - docling (opt-in; install `pip install docling` — much better PDF/table fidelity)
  - pandoc (opt-in; needs `pandoc` on PATH — best round-trip safety, esp. DOCX↔MD)

Security:
  - Local files only by default; absolute paths required.
  - Optional workspace allow-list via MARKITDOWN_DIFF_ALLOWED_ROOTS (':'-separated dirs).
  - File size cap (default 25 MB) and output size cap (default 5 MB).

JSON output schema (per spec, jsondiffpatch-inspired hunks):
  {summary, inputs, diffs:[{type, before_start, before_end, after_start, after_end,
                            before, after, tokens:[{type, text}]}],
   unified_diff, markdown_report?}
"""

from __future__ import annotations

import difflib
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Annotated, Any, Dict, List, Literal, Optional, Tuple

from fastmcp import FastMCP
from pydantic import BeforeValidator

mcp = FastMCP("markitdown-diff")


# ---------------------------------------------------------------------------
# Type coercion — accept dict/list params as either real JSON or a JSON-encoded
# string. Some MCP clients string-serialize structured args before sending; this
# keeps the tool usable from every client without losing strict validation.
# ---------------------------------------------------------------------------
def _coerce_dict_arg(v: Any) -> Any:
    if v is None or isinstance(v, dict):
        return v
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            parsed = json.loads(s)
        except json.JSONDecodeError as e:
            raise ValueError(f"expected a JSON object; got unparseable string: {e}")
        if not isinstance(parsed, dict):
            raise ValueError("expected a JSON object")
        return parsed
    raise ValueError(f"expected an object/dict, got {type(v).__name__}")


def _coerce_list_arg(v: Any) -> Any:
    if v is None or isinstance(v, list):
        return v
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            parsed = json.loads(s)
        except json.JSONDecodeError as e:
            raise ValueError(f"expected a JSON array; got unparseable string: {e}")
        if not isinstance(parsed, list):
            raise ValueError("expected a JSON array")
        return parsed
    raise ValueError(f"expected an array/list, got {type(v).__name__}")


_Obj = Annotated[Optional[Dict[str, Any]], BeforeValidator(_coerce_dict_arg)]
_List = Annotated[Optional[List[Any]], BeforeValidator(_coerce_list_arg)]

# ---------------------------------------------------------------------------
# Config + security
# ---------------------------------------------------------------------------
DEFAULT_CONVERTER = os.environ.get("MARKITDOWN_DIFF_DEFAULT_CONVERTER", "markitdown").lower()
MAX_FILE_BYTES = int(os.environ.get("MARKITDOWN_DIFF_MAX_FILE_BYTES", str(25 * 1024 * 1024)))
MAX_OUTPUT_CHARS = int(os.environ.get("MARKITDOWN_DIFF_MAX_OUTPUT_CHARS", str(5 * 1024 * 1024)))
_ALLOWED_ROOTS_RAW = os.environ.get("MARKITDOWN_DIFF_ALLOWED_ROOTS", "")
_ALLOWED_ROOTS: List[pathlib.Path] = (
    [pathlib.Path(p).expanduser().resolve() for p in _ALLOWED_ROOTS_RAW.split(":") if p.strip()]
    if _ALLOWED_ROOTS_RAW
    else []
)

DIRECT_READ_EXT = {".md", ".markdown", ".txt", ".rst", ".csv", ".json", ".xml", ".html", ".htm"}
CONVERT_EXT = {".docx", ".pdf", ".pptx", ".xlsx", ".xls", ".epub"}

ErrorCode = Literal[
    "INVALID_PATH", "PATH_OUTSIDE_ALLOWED_ROOT", "FILE_TOO_LARGE",
    "UNSUPPORTED_FORMAT", "CONVERSION_FAILED", "DIFF_FAILED", "OUTPUT_TOO_LARGE",
    "CONVERTER_UNAVAILABLE", "INVALID_INPUT",
]


def _err(code: ErrorCode, message: str, **details: Any) -> Dict[str, Any]:
    return {"error": {"code": code, "message": message, "details": details}}


def _resolve_path(path_str: str) -> pathlib.Path:
    if not path_str or not isinstance(path_str, str):
        raise ValueError(("INVALID_PATH", "path must be a non-empty string"))
    p = pathlib.Path(path_str)
    if not p.is_absolute():
        raise ValueError(("INVALID_PATH", f"path must be absolute: {path_str!r}"))
    try:
        p = p.resolve(strict=True)
    except FileNotFoundError as e:
        raise ValueError(("INVALID_PATH", f"file not found: {path_str!r}")) from e
    if _ALLOWED_ROOTS:
        if not any(str(p).startswith(str(root) + os.sep) or p == root for root in _ALLOWED_ROOTS):
            raise ValueError(("PATH_OUTSIDE_ALLOWED_ROOT",
                              f"path is outside MARKITDOWN_DIFF_ALLOWED_ROOTS: {p}"))
    size = p.stat().st_size
    if size > MAX_FILE_BYTES:
        raise ValueError(("FILE_TOO_LARGE", f"file is {size} bytes; cap is {MAX_FILE_BYTES}"))
    return p


# ---------------------------------------------------------------------------
# Converter dispatcher
# ---------------------------------------------------------------------------
def _convert_with_markitdown(path: pathlib.Path) -> str:
    try:
        from markitdown import MarkItDown  # type: ignore
    except ImportError as e:
        raise ValueError(("CONVERTER_UNAVAILABLE", "markitdown not installed (pip install 'markitdown[all]')")) from e
    md = MarkItDown(enable_plugins=False)
    try:
        result = md.convert(str(path))
    except Exception as e:
        raise ValueError(("CONVERSION_FAILED", f"markitdown failed on {path.name}: {type(e).__name__}: {e}")) from e
    return getattr(result, "text_content", None) or getattr(result, "markdown", "") or ""


def _convert_with_docling(path: pathlib.Path) -> str:
    try:
        from docling.document_converter import DocumentConverter  # type: ignore
    except ImportError as e:
        raise ValueError(("CONVERTER_UNAVAILABLE", "docling not installed (pip install docling)")) from e
    try:
        result = DocumentConverter().convert(str(path))
        return result.document.export_to_markdown()
    except Exception as e:
        raise ValueError(("CONVERSION_FAILED", f"docling failed on {path.name}: {type(e).__name__}: {e}")) from e


def _convert_with_pandoc(path: pathlib.Path) -> str:
    try:
        import pypandoc  # type: ignore
    except ImportError as e:
        raise ValueError(("CONVERTER_UNAVAILABLE", "pypandoc not installed (pip install pypandoc) — also needs pandoc on PATH")) from e
    try:
        return pypandoc.convert_file(str(path), to="gfm")
    except Exception as e:
        raise ValueError(("CONVERSION_FAILED", f"pandoc failed on {path.name}: {type(e).__name__}: {e}")) from e


CONVERTERS = {
    "markitdown": _convert_with_markitdown,
    "docling": _convert_with_docling,
    "pandoc": _convert_with_pandoc,
}


def _convert(path: pathlib.Path, converter: str) -> str:
    fn = CONVERTERS.get(converter)
    if fn is None:
        raise ValueError(("INVALID_INPUT", f"unknown converter {converter!r}; choose from {list(CONVERTERS)}"))
    return fn(path)


def _read_text(path: pathlib.Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def _load_as_markdown(path: pathlib.Path, mode: str, converter: str) -> Tuple[str, str, bool]:
    """Return (text, format_label, converted_bool)."""
    ext = path.suffix.lower()
    if mode == "plain_text":
        return _read_text(path), ext.lstrip(".") or "txt", False
    if mode == "markdown":
        return _read_text(path), "md", False
    if mode == "convert_to_markdown":
        return _convert(path, converter), ext.lstrip(".") or "unknown", True
    # auto
    if ext in DIRECT_READ_EXT:
        return _read_text(path), ext.lstrip("."), False
    if ext in CONVERT_EXT:
        return _convert(path, converter), ext.lstrip("."), True
    raise ValueError(("UNSUPPORTED_FORMAT", f"unsupported extension: {ext or '(none)'}", ))


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------
_FRONTMATTER_RE = re.compile(r"\A---\n.*?\n---\n?", re.DOTALL)


def _normalize(text: str, opts: Optional[Dict[str, Any]] = None) -> str:
    o = dict(opts or {})
    # always-on safe defaults
    out = text.replace("\r\n", "\n").replace("\r", "\n")
    if o.get("trim_trailing_whitespace", True):
        out = "\n".join(line.rstrip() for line in out.split("\n"))
    # opt-in
    if o.get("ignore_frontmatter"):
        out = _FRONTMATTER_RE.sub("", out, count=1)
    if o.get("collapse_blank_lines"):
        out = re.sub(r"\n{3,}", "\n\n", out)
    if o.get("ignore_blank_lines"):
        out = re.sub(r"\n\s*\n+", "\n", out)
    if o.get("ignore_whitespace"):
        # collapse runs of whitespace inside lines to a single space
        out = "\n".join(re.sub(r"\s+", " ", ln).strip() for ln in out.split("\n"))
    if o.get("ignore_case"):
        out = out.lower()
    if o.get("normalize_unicode_nfc"):
        import unicodedata
        out = unicodedata.normalize("NFC", out)
    if o.get("remove_html_comments"):
        out = re.sub(r"<!--.*?-->", "", out, flags=re.DOTALL)
    if o.get("normalize_smart_quotes"):
        out = (out.replace("“", '"').replace("”", '"')
                   .replace("‘", "'").replace("’", "'"))
    if not out.endswith("\n"):
        out += "\n"
    return out


# ---------------------------------------------------------------------------
# Text statistics — chars, words, sentences, paragraphs, reading time
# ---------------------------------------------------------------------------
_SENTENCE_END_RE = re.compile(r"[.!?][\)\]\"']*(?=\s|$)")
_WORD_COUNT_RE = re.compile(r"\b\w+\b", re.UNICODE)
_READING_WPM = 230  # average silent reading
_SPEAKING_WPM = 150


def _text_stats(text: str) -> Dict[str, Any]:
    chars_total = len(text)
    chars_no_spaces = len(re.sub(r"\s", "", text))
    letters = len(re.sub(r"[^\w]", "", text))
    words = _WORD_COUNT_RE.findall(text)
    word_count = len(words)
    lines = text.split("\n")
    line_count = len(lines)
    # paragraphs = blocks separated by blank lines (or single line if no blanks)
    paragraphs = [p for p in re.split(r"\n\s*\n+", text.strip()) if p.strip()]
    paragraph_count = len(paragraphs)
    # sentences = end-punct splits; fallback to paragraph count
    sentence_count = max(len(_SENTENCE_END_RE.findall(text)), 1 if text.strip() else 0)
    avg_word_len = round(sum(len(w) for w in words) / word_count, 2) if word_count else 0.0
    reading_min = round(word_count / _READING_WPM, 2) if word_count else 0.0
    speaking_min = round(word_count / _SPEAKING_WPM, 2) if word_count else 0.0
    return {
        "characters": chars_total,
        "characters_no_spaces": chars_no_spaces,
        "letters": letters,
        "words": word_count,
        "sentences": sentence_count,
        "paragraphs": paragraph_count,
        "lines": line_count,
        "avg_word_length": avg_word_len,
        "reading_time_minutes": reading_min,
        "speaking_time_minutes": speaking_min,
    }


# ---------------------------------------------------------------------------
# Diff engine (difflib)
# ---------------------------------------------------------------------------
_WORD_TOKEN_RE = re.compile(r"\w+|\s+|[^\w\s]", re.UNICODE)


def _tokenize_words(s: str) -> List[str]:
    return _WORD_TOKEN_RE.findall(s)


def _token_diff(before: str, after: str) -> List[Dict[str, str]]:
    b_tokens = _tokenize_words(before)
    a_tokens = _tokenize_words(after)
    sm = difflib.SequenceMatcher(a=b_tokens, b=a_tokens, autojunk=False)
    out: List[Dict[str, str]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            if i1 != i2:
                out.append({"type": "unchanged", "text": "".join(b_tokens[i1:i2])})
        elif tag == "delete":
            out.append({"type": "removed", "text": "".join(b_tokens[i1:i2])})
        elif tag == "insert":
            out.append({"type": "added", "text": "".join(a_tokens[j1:j2])})
        elif tag == "replace":
            out.append({"type": "removed", "text": "".join(b_tokens[i1:i2])})
            out.append({"type": "added", "text": "".join(a_tokens[j1:j2])})
    return out


_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+")


def _split_units(text: str, granularity: str) -> List[str]:
    if granularity == "sentence":
        # split by sentence-ending punctuation followed by whitespace;
        # collapse newlines within sentences to spaces for cleaner diffs
        flat = re.sub(r"\s+", " ", text).strip()
        return _SENTENCE_SPLIT_RE.split(flat) if flat else []
    # default: lines
    return text.split("\n")


def _extract_replacements(diffs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Walk modified hunks' token diffs, group adjacent removed->added pairs
    and count their frequency across the document. The classic DiffChecker
    'what got replaced' view."""
    pairs: Dict[Tuple[str, str], int] = {}
    for d in diffs:
        toks = d.get("tokens")
        if not toks or d.get("type") != "modified":
            continue
        i = 0
        while i < len(toks):
            if toks[i]["type"] == "removed":
                # collapse adjacent removed tokens into one phrase
                rem_text = toks[i]["text"]
                j = i + 1
                while j < len(toks) and toks[j]["type"] == "removed":
                    rem_text += toks[j]["text"]
                    j += 1
                # collapse adjacent added tokens that follow
                add_text = ""
                while j < len(toks) and toks[j]["type"] == "added":
                    add_text += toks[j]["text"]
                    j += 1
                rem_norm = rem_text.strip()
                add_norm = add_text.strip()
                if rem_norm and add_norm:
                    pairs[(rem_norm, add_norm)] = pairs.get((rem_norm, add_norm), 0) + 1
                i = j
            else:
                i += 1
    return [
        {"removed": r, "added": a, "count": c}
        for (r, a), c in sorted(pairs.items(), key=lambda kv: -kv[1])
    ]


def _line_diff(before: str, after: str, granularity: str,
               hide_unchanged: bool = False) -> Dict[str, Any]:
    unit_gran = "sentence" if granularity == "sentence" else "line"
    b_units = _split_units(before, unit_gran)
    a_units = _split_units(after, unit_gran)
    joiner = " " if unit_gran == "sentence" else "\n"
    sm = difflib.SequenceMatcher(a=b_units, b=a_units, autojunk=False)
    diffs: List[Dict[str, Any]] = []
    added = removed = modified = unchanged = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            unchanged += i2 - i1
            if not hide_unchanged:
                # only emit unchanged hunks if the caller wants them; skipping
                # is the DiffChecker-style default for compact JSON.
                pass
            continue
        before_chunk = joiner.join(b_units[i1:i2])
        after_chunk = joiner.join(a_units[j1:j2])
        if tag == "insert":
            added += j2 - j1
            diffs.append({
                "type": "added",
                "before_start": i1 + 1, "before_end": i1,
                "after_start": j1 + 1, "after_end": j2,
                "before": "", "after": after_chunk,
            })
        elif tag == "delete":
            removed += i2 - i1
            diffs.append({
                "type": "removed",
                "before_start": i1 + 1, "before_end": i2,
                "after_start": j1 + 1, "after_end": j1,
                "before": before_chunk, "after": "",
            })
        else:  # replace
            modified += 1
            removed += i2 - i1
            added += j2 - j1
            entry: Dict[str, Any] = {
                "type": "modified",
                "before_start": i1 + 1, "before_end": i2,
                "after_start": j1 + 1, "after_end": j2,
                "before": before_chunk, "after": after_chunk,
            }
            if granularity in ("word", "char", "sentence"):
                entry["tokens"] = _token_diff(before_chunk, after_chunk)
            diffs.append(entry)
    ratio = sm.ratio()
    unit_label = "sentences" if unit_gran == "sentence" else "lines"
    return {
        "diffs": diffs,
        "summary": {
            f"added_{unit_label}": added,
            f"removed_{unit_label}": removed,
            "modified_hunks": modified,
            f"unchanged_{unit_label}": unchanged,
            "similarity": round(ratio, 6),
        },
    }


def _unified_diff(before: str, after: str, context_lines: int,
                  fromfile: str = "before", tofile: str = "after") -> str:
    return "".join(difflib.unified_diff(
        before.splitlines(keepends=True),
        after.splitlines(keepends=True),
        fromfile=fromfile, tofile=tofile, n=context_lines,
    ))


def _markdown_report(result: Dict[str, Any]) -> str:
    s = result["summary"]
    lines = [
        "## Diff Summary",
        "",
        f"- Lines added: **{s['added_lines']}**",
        f"- Lines removed: **{s['removed_lines']}**",
        f"- Modified hunks: **{s['modified_hunks']}**",
        f"- Unchanged lines: **{s['unchanged_lines']}**",
        f"- Similarity: **{s['similarity']:.3f}**",
        "",
    ]
    if not result["diffs"]:
        lines.append("_No differences._")
        return "\n".join(lines)
    lines.append("## Changes")
    lines.append("")
    for d in result["diffs"]:
        marker = {"added": "+", "removed": "-", "modified": "~"}.get(d["type"], "?")
        lines.append(f"### {marker} `{d['type']}` — before L{d['before_start']}-{d['before_end']}, after L{d['after_start']}-{d['after_end']}")
        lines.append("")
        if d["type"] == "modified" and "tokens" in d:
            inline_parts: List[str] = []
            for tok in d["tokens"]:
                if tok["type"] == "removed":
                    inline_parts.append(f"~~{tok['text']}~~")
                elif tok["type"] == "added":
                    inline_parts.append(f"**{tok['text']}**")
                else:
                    inline_parts.append(tok["text"])
            lines.append("> " + "".join(inline_parts).replace("\n", "\n> "))
        else:
            if d.get("before"):
                lines.append("```diff")
                for ln in d["before"].split("\n"):
                    lines.append(f"- {ln}")
                lines.append("```")
            if d.get("after"):
                lines.append("```diff")
                for ln in d["after"].split("\n"):
                    lines.append(f"+ {ln}")
                lines.append("```")
        lines.append("")
    return "\n".join(lines)


def _build_result(
    before: str, after: str, *,
    granularity: str, output: str, context_lines: int,
    before_meta: Dict[str, Any], after_meta: Dict[str, Any],
    hide_unchanged: bool = False, include_stats: bool = True,
    include_replacements: bool = True,
) -> Dict[str, Any]:
    core = _line_diff(before, after, granularity, hide_unchanged=hide_unchanged)
    b_stats = _text_stats(before) if include_stats else None
    a_stats = _text_stats(after) if include_stats else None

    # Per-side counts folded into the summary (the DiffChecker-style headline)
    summary: Dict[str, Any] = dict(core["summary"])
    if b_stats and a_stats:
        summary["before_words"] = b_stats["words"]
        summary["after_words"] = a_stats["words"]
        summary["words_delta"] = a_stats["words"] - b_stats["words"]
        summary["before_characters"] = b_stats["characters"]
        summary["after_characters"] = a_stats["characters"]
        summary["characters_delta"] = a_stats["characters"] - b_stats["characters"]
        summary["before_sentences"] = b_stats["sentences"]
        summary["after_sentences"] = a_stats["sentences"]
        summary["before_paragraphs"] = b_stats["paragraphs"]
        summary["after_paragraphs"] = a_stats["paragraphs"]

    res: Dict[str, Any] = {
        "summary": summary,
        "inputs": {
            "before": {**before_meta, "markdown_chars": len(before),
                       **({"stats": b_stats} if b_stats else {})},
            "after": {**after_meta, "markdown_chars": len(after),
                      **({"stats": a_stats} if a_stats else {})},
        },
    }
    if include_replacements:
        res["replacements"] = _extract_replacements(core["diffs"])
        summary["replacements"] = len(res["replacements"])
    if output in ("json", "all"):
        res["diffs"] = core["diffs"]
    if output in ("unified", "all"):
        res["unified_diff"] = _unified_diff(before, after, context_lines)
    if output in ("markdown", "all"):
        res["markdown_report"] = _markdown_report(core)
    # output cap
    blob = json.dumps(res, ensure_ascii=False)
    if len(blob) > MAX_OUTPUT_CHARS:
        return _err("OUTPUT_TOO_LARGE",
                    f"output is {len(blob)} chars; cap is {MAX_OUTPUT_CHARS}. "
                    "Lower granularity, drop the markdown_report, hide_unchanged, "
                    "or raise MARKITDOWN_DIFF_MAX_OUTPUT_CHARS.")
    return res


# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------
@mcp.tool(
    description=(
        "Compare two text strings (no file access). Returns structured JSON "
        "hunks PLUS per-side stats (words, characters, sentences, paragraphs), "
        "a replacements roll-up (the 'may→shall x3' view DiffChecker is famous "
        "for), and optional unified-diff / Markdown report. Args: before, after; "
        "granularity ('line'|'word'|'char'|'sentence', default 'line'); output "
        "('json'|'unified'|'markdown'|'all', default 'json'); context_lines "
        "(default 3); hide_unchanged (default False — when True, omits unchanged "
        "regions for compact output)."
    ),
)
async def compare_text(
    before: str, after: str,
    granularity: str = "line", output: str = "json", context_lines: int = 3,
    hide_unchanged: bool = False,
) -> Dict[str, Any]:
    try:
        if not isinstance(before, str) or not isinstance(after, str):
            return _err("INVALID_INPUT", "before and after must be strings")
        return _build_result(
            before, after,
            granularity=granularity, output=output, context_lines=context_lines,
            hide_unchanged=hide_unchanged,
            before_meta={"source": "inline", "format": "text", "converted": False},
            after_meta={"source": "inline", "format": "text", "converted": False},
        )
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Compare two Markdown strings with Markdown-aware normalization. Args: "
        "before, after; normalize (dict; flags: trim_trailing_whitespace, "
        "collapse_blank_lines, ignore_blank_lines, ignore_whitespace, "
        "ignore_case, ignore_frontmatter, normalize_unicode_nfc, "
        "remove_html_comments, normalize_smart_quotes); granularity ('line'|"
        "'word'|'char'|'sentence'); output ('json'|'unified'|'markdown'|'all'); "
        "context_lines; hide_unchanged. Returns the same rich shape as "
        "compare_text — per-side stats, replacements roll-up, hunks, and reports."
    ),
)
async def compare_markdown(
    before: str, after: str,
    normalize: _Obj = None,
    granularity: str = "word", output: str = "json", context_lines: int = 3,
    hide_unchanged: bool = False,
) -> Dict[str, Any]:
    try:
        if not isinstance(before, str) or not isinstance(after, str):
            return _err("INVALID_INPUT", "before and after must be strings")
        b = _normalize(before, normalize)
        a = _normalize(after, normalize)
        return _build_result(
            b, a,
            granularity=granularity, output=output, context_lines=context_lines,
            hide_unchanged=hide_unchanged,
            before_meta={"source": "inline", "format": "md", "converted": False},
            after_meta={"source": "inline", "format": "md", "converted": False},
        )
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Compare two LOCAL files. Markdown and text files are read directly; "
        "DOCX/PDF/PPTX/XLSX/EPUB are converted to Markdown first. Args: "
        "before_path, after_path (absolute); mode ('auto'|'plain_text'|'markdown'|"
        "'convert_to_markdown'); converter ('markitdown'|'docling'|'pandoc' — "
        "default markitdown; docling recommended for PDF fidelity); granularity; "
        "output; context_lines; normalize (dict). Returns JSON/unified/markdown "
        "report plus input metadata indicating which converter ran."
    ),
)
async def compare_files(
    before_path: str, after_path: str,
    mode: str = "auto", converter: Optional[str] = None,
    granularity: str = "word", output: str = "json", context_lines: int = 3,
    normalize: _Obj = None,
    hide_unchanged: bool = False,
) -> Dict[str, Any]:
    try:
        try:
            b_path = _resolve_path(before_path)
            a_path = _resolve_path(after_path)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        conv = (converter or DEFAULT_CONVERTER).lower()
        try:
            b_text, b_fmt, b_converted = _load_as_markdown(b_path, mode, conv)
            a_text, a_fmt, a_converted = _load_as_markdown(a_path, mode, conv)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        b_text = _normalize(b_text, normalize)
        a_text = _normalize(a_text, normalize)
        return _build_result(
            b_text, a_text,
            granularity=granularity, output=output, context_lines=context_lines,
            hide_unchanged=hide_unchanged,
            before_meta={"source": str(b_path), "format": b_fmt, "converted": b_converted,
                         "converter": conv if b_converted else None},
            after_meta={"source": str(a_path), "format": a_fmt, "converted": a_converted,
                        "converter": conv if a_converted else None},
        )
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


# ---------------------------------------------------------------------------
# CSS — structural-aware diff via Prettier-normalize + difftastic
# ---------------------------------------------------------------------------
# CSS diff has four problems that text diff doesn't handle: whitespace/format
# noise, rule-order shuffles, selector-grouping equivalences, and per-rule
# property order. Pre-formatting both sides with Prettier collapses ~90% of
# that noise into a canonical shape, which the existing line-diff engine then
# handles correctly. `difft` (difftastic) is an optional second view —
# tree-sitter-backed structural diff rendered as terminal output, useful when
# a human (rather than an agent) is reviewing the result.

_CSS_PARSERS = {"css", "scss", "less"}


def _prettier_normalize_css(text: str, parser: str = "css") -> str:
    """Pipe text through `prettier --parser <parser>` for canonical formatting.
    Returns the input unchanged if prettier isn't on PATH or fails — best-effort
    so the tool stays useful when prettier isn't installed."""
    bin_path = shutil.which("prettier")
    if not bin_path:
        return text
    try:
        r = subprocess.run(
            [bin_path, "--parser", parser, "--stdin-filepath", f"input.{parser}"],
            input=text, capture_output=True, text=True, timeout=30,
        )
        if r.returncode == 0 and r.stdout:
            return r.stdout
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        pass
    return text


def _run_difftastic(before: str, after: str, suffix: str = ".css") -> Optional[str]:
    """Run difft against two tempfiles, return its stdout (no color codes).
    Returns None if difft isn't installed or the run fails — tool stays useful
    without it."""
    bin_path = shutil.which("difft")
    if not bin_path:
        return None
    b_path = a_path = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=suffix, delete=False, encoding="utf-8") as f:
            f.write(before); b_path = f.name
        with tempfile.NamedTemporaryFile(mode="w", suffix=suffix, delete=False, encoding="utf-8") as f:
            f.write(after); a_path = f.name
        r = subprocess.run(
            [bin_path, "--color=never", b_path, a_path],
            capture_output=True, text=True, timeout=30,
        )
        return r.stdout or None
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return None
    finally:
        for p in (b_path, a_path):
            if p:
                try:
                    os.unlink(p)
                except OSError:
                    pass


@mcp.tool(
    description=(
        "Compare two CSS / SCSS / LESS stylesheets — the dedicated CSS-aware "
        "diff tool. Use whenever you need to: see what changed in a "
        "stylesheet, diff a developer's CSS update, compare WordPress theme "
        "files across sites, audit which selectors / properties / CSS custom "
        "properties (variables) changed between versions, review a CSS pull "
        "request, compare before/after of a build pipeline, or check what's "
        "different between two stylesheets. Pre-formats both sides with "
        "Prettier first so formatting noise (whitespace, brace style, property "
        "spacing, minified-vs-prettified, property ordering within a rule) "
        "doesn't drown out the real semantic changes — then runs the structured "
        "line-diff engine to surface added / removed / modified rules and "
        "declarations. Much better than raw text diff or `git diff` for CSS "
        "specifically.\n"
        "Accepts either inline `before` + `after` strings OR `before_path` + "
        "`after_path` file paths (absolute). Returns the same rich JSON shape "
        "as compare_files: per-side stats (chars/words/lines/sentences), hunks "
        "with line numbers + word-level token edits inside modified blocks, "
        "replacements roll-up (the 'this rule → that rule' view), optional "
        "unified diff and Markdown report. Set `include_difftastic=True` for "
        "an additional syntax-aware tree-edit view from `difft` (best when a "
        "human is reviewing — agents will usually prefer the JSON hunks).\n"
        "Args: before / after (strings) OR before_path / after_path (absolute "
        "paths); normalize (default True — set False to skip Prettier and "
        "diff raw bytes); parser ('css' default | 'scss' | 'less'); "
        "granularity ('line' default — word-level is usually noise on CSS); "
        "output ('json' | 'unified' | 'markdown' | 'all'); context_lines; "
        "hide_unchanged; include_difftastic. Prettier and difft are auto-"
        "detected; the tool degrades gracefully if either is missing — the "
        "diff still runs against unnormalized text."
    ),
)
async def compare_css(
    before: Optional[str] = None, after: Optional[str] = None,
    before_path: Optional[str] = None, after_path: Optional[str] = None,
    normalize: bool = True, parser: str = "css",
    granularity: str = "line", output: str = "json", context_lines: int = 3,
    hide_unchanged: bool = False, include_difftastic: bool = False,
) -> Dict[str, Any]:
    try:
        if parser not in _CSS_PARSERS:
            return _err("INVALID_INPUT", f"parser must be one of {sorted(_CSS_PARSERS)}; got {parser!r}")

        def _load_side(text: Optional[str], p: Optional[str], label: str) -> Tuple[str, str, str]:
            if text is not None and p is not None:
                raise ValueError(("INVALID_INPUT", f"provide either {label} or {label}_path, not both"))
            if text is not None:
                if not isinstance(text, str):
                    raise ValueError(("INVALID_INPUT", f"{label} must be a string"))
                return text, "inline", parser
            if p is None:
                raise ValueError(("INVALID_INPUT", f"provide either {label} or {label}_path"))
            path = _resolve_path(p)
            return path.read_text(encoding="utf-8", errors="replace"), str(path), (path.suffix.lstrip(".") or parser)

        try:
            b_text, b_src, b_fmt = _load_side(before, before_path, "before")
            a_text, a_src, a_fmt = _load_side(after, after_path, "after")
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)

        normalized_flag = False
        if normalize:
            b_norm = _prettier_normalize_css(b_text, parser)
            a_norm = _prettier_normalize_css(a_text, parser)
            normalized_flag = bool(shutil.which("prettier"))
            b_text, a_text = b_norm, a_norm

        result = _build_result(
            b_text, a_text,
            granularity=granularity, output=output, context_lines=context_lines,
            hide_unchanged=hide_unchanged,
            before_meta={"source": b_src, "format": b_fmt, "converted": False,
                         "normalized": normalized_flag, "parser": parser},
            after_meta={"source": a_src, "format": a_fmt, "converted": False,
                        "normalized": normalized_flag, "parser": parser},
        )

        if include_difftastic and not (isinstance(result, dict) and "error" in result):
            suffix = "." + parser
            dft = _run_difftastic(b_text, a_text, suffix=suffix)
            if dft:
                result["difftastic_view"] = dft

        return result
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Compute comprehensive text statistics for a single string — the kind "
        "of counts DiffChecker shows in its sidebar. Returns characters (with "
        "and without spaces), letters, words, sentences, paragraphs, lines, "
        "average word length, and estimated reading + speaking time in minutes. "
        "Useful standalone (for character-limit checks, length budgets) or as a "
        "lightweight prelude to compare_* tools."
    ),
)
async def text_stats(text: str) -> Dict[str, Any]:
    try:
        if not isinstance(text, str):
            return _err("INVALID_INPUT", "text must be a string")
        return _text_stats(text)
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Convert a local file to Markdown via the chosen converter. Args: path "
        "(absolute); output_path (optional — if provided, also writes Markdown "
        "to disk); converter ('markitdown'|'docling'|'pandoc'); include_metadata "
        "(reserved). Returns {markdown, format, converter, source, output_path}."
    ),
)
async def normalize_to_markdown(
    path: str, output_path: Optional[str] = None,
    converter: Optional[str] = None, include_metadata: bool = False,
) -> Dict[str, Any]:
    try:
        try:
            p = _resolve_path(path)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        conv = (converter or DEFAULT_CONVERTER).lower()
        ext = p.suffix.lower()
        try:
            if ext in DIRECT_READ_EXT:
                md_text = _read_text(p)
                converted = False
            elif ext in CONVERT_EXT:
                md_text = _convert(p, conv)
                converted = True
            else:
                return _err("UNSUPPORTED_FORMAT", f"unsupported extension: {ext or '(none)'}")
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        written: Optional[str] = None
        if output_path:
            try:
                op = pathlib.Path(output_path)
                if not op.is_absolute():
                    return _err("INVALID_PATH", f"output_path must be absolute: {output_path!r}")
                op.parent.mkdir(parents=True, exist_ok=True)
                op.write_text(md_text, encoding="utf-8")
                written = str(op)
            except OSError as e:
                return _err("CONVERSION_FAILED", f"failed to write output: {e}")
        return {
            "markdown": md_text,
            "format": ext.lstrip("."),
            "converter": conv if converted else None,
            "converted": converted,
            "source": str(p),
            "output_path": written,
        }
    except Exception as e:
        return _err("CONVERSION_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Diff two strings AND return a short natural-language summary alongside "
        "the structured diff — useful for agent-driven document review. Args: "
        "before, after; normalize (dict, applied if both strings look like "
        "Markdown); granularity ('word' default). Returns {summary_text, summary, "
        "diffs, similarity}."
    ),
)
async def summarize_diff(
    before: str, after: str,
    normalize: _Obj = None, granularity: str = "word",
) -> Dict[str, Any]:
    try:
        if not isinstance(before, str) or not isinstance(after, str):
            return _err("INVALID_INPUT", "before and after must be strings")
        b = _normalize(before, normalize)
        a = _normalize(after, normalize)
        core = _line_diff(b, a, granularity)
        s = core["summary"]
        if not core["diffs"]:
            text = "No differences."
        else:
            parts = []
            if s["added_lines"]:
                parts.append(f"{s['added_lines']} line(s) added")
            if s["removed_lines"]:
                parts.append(f"{s['removed_lines']} line(s) removed")
            if s["modified_hunks"]:
                parts.append(f"{s['modified_hunks']} hunk(s) modified")
            text = "; ".join(parts) + f". Similarity: {s['similarity']:.1%}."
        return {
            "summary_text": text,
            "summary": s,
            "diffs": core["diffs"],
            "similarity": s["similarity"],
        }
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


# ---------------------------------------------------------------------------
# Find / replace — preview-first, optional apply
# ---------------------------------------------------------------------------
def _find_matches(text: str, pattern: str, regex: bool = False,
                  max_matches: int = 200) -> List[Dict[str, Any]]:
    """Return [{line, column, match, line_text}, ...] for each pattern hit."""
    if not pattern:
        return []
    spans: List[Tuple[int, str]] = []
    if regex:
        try:
            pat = re.compile(pattern, re.MULTILINE)
        except re.error as e:
            raise ValueError(("INVALID_INPUT", f"invalid regex: {e}"))
        for m in pat.finditer(text):
            spans.append((m.start(), m.group()))
            if len(spans) >= max_matches:
                break
    else:
        i = 0
        while True:
            j = text.find(pattern, i)
            if j < 0:
                break
            spans.append((j, pattern))
            if len(spans) >= max_matches:
                break
            i = j + max(1, len(pattern))
    out: List[Dict[str, Any]] = []
    for start, matched in spans:
        prefix = text[:start]
        line_no = prefix.count("\n") + 1
        line_start = prefix.rfind("\n") + 1
        col_no = start - line_start + 1
        line_end = text.find("\n", start)
        if line_end < 0:
            line_end = len(text)
        out.append({
            "line": line_no, "column": col_no,
            "match": matched, "line_text": text[line_start:line_end],
        })
    return out


def _normalize_operations(operations: Optional[List[Dict[str, Any]]],
                          pattern: Optional[str], replacement: str,
                          regex: bool) -> List[Dict[str, Any]]:
    if operations:
        if not isinstance(operations, list) or not all(isinstance(o, dict) for o in operations):
            raise ValueError(("INVALID_INPUT", "operations must be a list of {pattern, replacement, regex?} objects"))
        out: List[Dict[str, Any]] = []
        for op in operations:
            p = op.get("pattern")
            if not p or not isinstance(p, str):
                raise ValueError(("INVALID_INPUT", "every operation needs a non-empty `pattern` string"))
            out.append({
                "pattern": p,
                "replacement": op.get("replacement", ""),
                "regex": bool(op.get("regex", regex)),
            })
        return out
    if not pattern:
        raise ValueError(("INVALID_INPUT", "provide `pattern` or `operations`"))
    return [{"pattern": pattern, "replacement": replacement, "regex": regex}]


def _preview_ops(text: str, ops: List[Dict[str, Any]],
                 max_matches: int = 200, inline_cap: int = 50) -> List[Dict[str, Any]]:
    previews: List[Dict[str, Any]] = []
    for op in ops:
        try:
            matches = _find_matches(text, op["pattern"], op["regex"], max_matches)
        except ValueError as e:
            raise e
        previews.append({
            "pattern": op["pattern"],
            "replacement": op["replacement"],
            "regex": op["regex"],
            "match_count": len(matches),
            "matches": matches[:inline_cap],
            "matches_truncated": len(matches) > inline_cap,
        })
    return previews


def _apply_ops(text: str, ops: List[Dict[str, Any]]) -> Tuple[str, List[Dict[str, Any]]]:
    out = text
    applied: List[Dict[str, Any]] = []
    for op in ops:
        if op["regex"]:
            try:
                new_text, n = re.subn(op["pattern"], op["replacement"], out)
            except re.error as e:
                raise ValueError(("INVALID_INPUT", f"invalid regex: {e}"))
        else:
            n = out.count(op["pattern"])
            new_text = out.replace(op["pattern"], op["replacement"])
        applied.append({
            "pattern": op["pattern"],
            "replacement": op["replacement"],
            "regex": op["regex"],
            "replacements_applied": n,
        })
        out = new_text
    return out, applied


@mcp.tool(
    description=(
        "Compute comprehensive text statistics for a LOCAL FILE. Converts "
        "DOCX/PDF/PPTX/XLSX/EPUB to Markdown first via the chosen converter, "
        "then runs the same counts as text_stats (characters, words, sentences, "
        "paragraphs, reading time, etc.). Args: path (absolute); converter "
        "('markitdown'|'docling'|'pandoc')."
    ),
)
async def count_file(path: str, converter: Optional[str] = None) -> Dict[str, Any]:
    try:
        try:
            p = _resolve_path(path)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        conv = (converter or DEFAULT_CONVERTER).lower()
        ext = p.suffix.lower()
        try:
            if ext in DIRECT_READ_EXT:
                text = _read_text(p)
                converted = False
            elif ext in CONVERT_EXT:
                text = _convert(p, conv)
                converted = True
            else:
                return _err("UNSUPPORTED_FORMAT", f"unsupported extension: {ext or '(none)'}")
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        return {
            "source": str(p),
            "format": ext.lstrip("."),
            "converted": converted,
            "converter": conv if converted else None,
            "stats": _text_stats(text),
        }
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Find-and-replace on a STRING with preview-first behavior. Defaults to "
        "DRY-RUN — returns matches with line + column locations and the "
        "containing line, no mutation. Set dry_run=False to compute and return "
        "the result string. Args: text; either (pattern, replacement) OR "
        "operations=[{pattern, replacement, regex?}]; regex (default False); "
        "dry_run (default True); max_matches (default 200, safety cap on "
        "previewed matches per op)."
    ),
)
async def find_replace_text(
    text: str, pattern: Optional[str] = None, replacement: str = "",
    operations: _List = None,
    regex: bool = False, dry_run: bool = True, max_matches: int = 200,
) -> Dict[str, Any]:
    try:
        if not isinstance(text, str):
            return _err("INVALID_INPUT", "text must be a string")
        try:
            ops = _normalize_operations(operations, pattern, replacement, regex)
            previews = _preview_ops(text, ops, max_matches=max_matches)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        if dry_run:
            return {
                "dry_run": True,
                "operations": previews,
                "total_matches": sum(p["match_count"] for p in previews),
            }
        try:
            new_text, applied = _apply_ops(text, ops)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        return {
            "dry_run": False,
            "operations": applied,
            "total_replacements": sum(a["replacements_applied"] for a in applied),
            "result": new_text,
        }
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


@mcp.tool(
    description=(
        "Find-and-replace on a LOCAL FILE with preview-first behavior. For "
        "non-text formats (DOCX, PDF, etc.) the file is converted to Markdown "
        "via the chosen converter first; replacements run on the Markdown text. "
        "Cannot write back to DOCX/PDF — provide output_path to save the "
        "modified Markdown elsewhere. For direct-text files (.md, .txt, .csv, "
        ".json, ...) you can pass overwrite_in_place=True to replace the file. "
        "Defaults to DRY-RUN. Args: path (absolute); either (pattern, replacement) "
        "OR operations; regex; dry_run; output_path (absolute, optional); "
        "overwrite_in_place (default False, ignored for converted files); "
        "converter; max_matches."
    ),
)
async def replace_in_file(
    path: str,
    pattern: Optional[str] = None, replacement: str = "",
    operations: _List = None,
    regex: bool = False, dry_run: bool = True,
    output_path: Optional[str] = None, overwrite_in_place: bool = False,
    converter: Optional[str] = None, max_matches: int = 200,
) -> Dict[str, Any]:
    try:
        try:
            p = _resolve_path(path)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)
        conv = (converter or DEFAULT_CONVERTER).lower()
        ext = p.suffix.lower()
        try:
            if ext in DIRECT_READ_EXT:
                text = _read_text(p)
                converted = False
            elif ext in CONVERT_EXT:
                text = _convert(p, conv)
                converted = True
            else:
                return _err("UNSUPPORTED_FORMAT", f"unsupported extension: {ext or '(none)'}")
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)

        try:
            ops = _normalize_operations(operations, pattern, replacement, regex)
            previews = _preview_ops(text, ops, max_matches=max_matches)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)

        if dry_run:
            return {
                "dry_run": True,
                "source": str(p), "format": ext.lstrip("."),
                "converted": converted,
                "converter": conv if converted else None,
                "operations": previews,
                "total_matches": sum(pv["match_count"] for pv in previews),
            }

        try:
            new_text, applied = _apply_ops(text, ops)
        except ValueError as e:
            code, msg = e.args[0]
            return _err(code, msg)

        written: Optional[str] = None
        if output_path:
            try:
                op_path = pathlib.Path(output_path)
                if not op_path.is_absolute():
                    return _err("INVALID_PATH", f"output_path must be absolute: {output_path!r}")
                op_path.parent.mkdir(parents=True, exist_ok=True)
                op_path.write_text(new_text, encoding="utf-8")
                written = str(op_path)
            except OSError as e:
                return _err("CONVERSION_FAILED", f"failed to write output: {e}")
        elif overwrite_in_place:
            if converted:
                return _err("INVALID_INPUT",
                            "cannot overwrite in place — source was a DOCX/PDF/etc. "
                            "Converted text is one-way; provide output_path to save the "
                            "modified Markdown instead.")
            p.write_text(new_text, encoding="utf-8")
            written = str(p)
        else:
            return _err("INVALID_INPUT",
                        "applied=true requires either output_path or overwrite_in_place=True (direct-text files only).")

        return {
            "dry_run": False,
            "source": str(p), "format": ext.lstrip("."),
            "converted": converted,
            "converter": conv if converted else None,
            "output_path": written,
            "operations": applied,
            "total_replacements": sum(a["replacements_applied"] for a in applied),
        }
    except Exception as e:
        return _err("DIFF_FAILED", f"{type(e).__name__}: {e}")


# ---------------------------------------------------------------------------
# Entry point + minimal CLI
# ---------------------------------------------------------------------------
def _cli() -> bool:
    args = sys.argv[1:]
    if not args:
        return False
    if args[0] == "version":
        print("markitdown-diff 0.1.0")
        return True
    if args[0] == "compare-files" and len(args) >= 3:
        import asyncio
        out = asyncio.run(compare_files(args[1], args[2], output=args[3] if len(args) > 3 else "markdown"))
        print(json.dumps(out, indent=2) if isinstance(out, dict) and "markdown_report" not in out
              else out.get("markdown_report", json.dumps(out, indent=2)))
        return True
    return False


def main() -> None:
    # Dispatch to CLI if args were passed; otherwise run as an MCP stdio server.
    if _cli():
        return
    mcp.run()


if __name__ == "__main__":
    main()
