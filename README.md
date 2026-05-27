# markitdown-diff-mcp

> A local **Model Context Protocol (MCP)** server for comparing Markdown, plain text, DOCX, PDF, PPTX, XLSX, and EPUB documents — with structured JSON hunks, unified diffs, Markdown reports, comprehensive text statistics, replacement preview/apply, and a converter dispatcher that runs **Microsoft MarkItDown** (default) or **IBM Docling** / **pandoc** under the hood.

A DiffChecker-style document comparison toolkit for Claude Desktop, Claude Code, Cursor, and any other MCP-aware AI assistant.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE) [![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/) [![MCP](https://img.shields.io/badge/MCP-2024--11--05-blueviolet)](https://modelcontextprotocol.io/)

---

## Why this exists

There are MCP servers that convert documents to Markdown. There are MCP servers that do raw text diff. But nothing wires the two together with **structured JSON output**, **per-side word/character/sentence statistics**, **replacement extraction** (the "may → shall × 3" view), and a **preview-first find/replace** primitive. This is that.

It runs locally, has no remote API dependency by default, and follows the same DiffChecker-style conventions agents and humans already expect.

## Tools

| Tool | What it does |
| --- | --- |
| `compare_text` | Diff two strings → JSON hunks, unified diff, Markdown report, per-side stats, replacements roll-up. |
| `compare_markdown` | Same, with Markdown-aware normalization (frontmatter, smart-quotes, etc.). |
| `compare_files` | Diff two local files (.md/.txt/.docx/.pdf/.pptx/.xlsx/.epub). Non-text formats are converted to Markdown first via the chosen converter. |
| `normalize_to_markdown` | Convert a single file to Markdown (optionally written to disk). |
| `text_stats` | Counts for a single string: characters, words, sentences, paragraphs, lines, reading & speaking time. |
| `count_file` | Same as `text_stats`, but against a file path (converts DOCX/PDF first). |
| `find_replace_text` | Preview-first find-and-replace on a string. Defaults to `dry_run=true`. Supports regex and a list of operations. Returns line + column of every match. |
| `replace_in_file` | Preview-first find-and-replace on a file. For DOCX/PDF, content is converted to Markdown first; result must be written to `output_path` (cannot write back to DOCX/PDF). For direct-text files, `overwrite_in_place=True` is allowed. |
| `summarize_diff` | Diff plus a short natural-language summary of what changed — useful for agent-driven document review. |

## Why three converters

| Converter | Best for | Notes |
| --- | --- | --- |
| `markitdown` *(default)* | DOCX, PPTX, HTML, in-process speed | Microsoft's MarkItDown — fast, great on Office formats, **weak on complex PDFs** (heading hierarchy & tables suffer per independent benchmarks). |
| `docling` *(opt-in)* | PDFs, scientific papers, multi-column layouts, tables | IBM's Docling — much higher fidelity on PDF (`pip install docling`). |
| `pandoc` *(opt-in)* | Round-trip-safe DOCX ↔ Markdown | Requires `pandoc` on `PATH` (`brew install pandoc`) plus `pip install pypandoc`. |

Pick per call with `converter="docling"` etc., or set the default with `MARKITDOWN_DIFF_DEFAULT_CONVERTER`.

## JSON output shape

Diff tools return:

```json
{
  "summary": {
    "added_lines": 2, "removed_lines": 2, "modified_hunks": 1,
    "unchanged_lines": 1, "similarity": 0.333,
    "before_words": 9, "after_words": 11, "words_delta": 2,
    "before_characters": 38, "after_characters": 47, "characters_delta": 9,
    "before_sentences": 2, "after_sentences": 2,
    "before_paragraphs": 1, "after_paragraphs": 1,
    "replacements": 1
  },
  "inputs": {
    "before": { "source": "...", "format": "md", "converted": false, "stats": { ... } },
    "after":  { "source": "...", "format": "docx", "converted": true, "converter": "markitdown", "stats": { ... } }
  },
  "replacements": [
    { "removed": "cat", "added": "dog", "count": 1 }
  ],
  "diffs": [
    {
      "type": "modified",
      "before_start": 1, "before_end": 1, "after_start": 1, "after_end": 1,
      "before": "The cat sat on the mat.",
      "after":  "The dog sat on the mat.",
      "tokens": [
        { "type": "unchanged", "text": "The " },
        { "type": "removed",   "text": "cat" },
        { "type": "added",     "text": "dog" },
        { "type": "unchanged", "text": " sat on the mat." }
      ]
    }
  ],
  "unified_diff": "@@ ... @@\n-The cat ...\n+The dog ...",
  "markdown_report": "## Diff Summary\n..."
}
```

The schema is jsondiffpatch-inspired so downstream tools can interop without translation.

## Install

```bash
# Base install (MarkItDown converter only)
pip install markitdown-diff-mcp

# With Docling for high-fidelity PDF
pip install 'markitdown-diff-mcp[docling]'

# With pandoc support
pip install 'markitdown-diff-mcp[pandoc]'

# Everything
pip install 'markitdown-diff-mcp[all]'
```

## Run

```bash
markitdown-diff-mcp        # stdio MCP server (for Claude Desktop, etc.)
```

Or run the single-file server directly:

```bash
uv run --script server.py  # uses the inline script header
```

## Configure in Claude Desktop

```jsonc
{
  "mcpServers": {
    "markitdown-diff": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/justinritchie/markitdown-diff-mcp", "markitdown-diff-mcp"],
      "env": {
        "MARKITDOWN_DIFF_DEFAULT_CONVERTER": "markitdown",
        "MARKITDOWN_DIFF_ALLOWED_ROOTS": "/Users/you/Documents:/Users/you/work"
      }
    }
  }
}
```

## Security

- Local files only by default; absolute paths required.
- Optional workspace allow-list via `MARKITDOWN_DIFF_ALLOWED_ROOTS` (colon-separated dirs). When set, any path outside is rejected with `PATH_OUTSIDE_ALLOWED_ROOT`.
- File size cap default 25 MB (`MARKITDOWN_DIFF_MAX_FILE_BYTES`).
- Output character cap default 5 MB (`MARKITDOWN_DIFF_MAX_OUTPUT_CHARS`).
- Find/replace defaults to `dry_run=true`. Writes require explicit opt-in.

## Roadmap

- v0.2: folder/directory compare, moved-block detection, HTML side-by-side report, Markdown AST-aware diffs (`markdown-it-py` + GFM tables).
- v0.3: optional integration with `docx-mcp` / `safe-docx` for true Word-track-changes export.
- Possible later: spreadsheet cell-by-cell diff via openpyxl, image visual diff (out of scope for v1; complex).

## Inspired by

- [Microsoft MarkItDown](https://github.com/microsoft/markitdown) — conversion engine
- [IBM Docling](https://github.com/DS4SD/docling) — high-fidelity PDF/layout parser
- [DiffChecker](https://www.diffchecker.com/) — UX conventions for stats, replacements, and preview-first behavior
- [Model Context Protocol](https://modelcontextprotocol.io/) — Anthropic + community spec

## License

MIT — see [LICENSE](LICENSE).
