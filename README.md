# html2docx

> Convert HTML to Word documents with CSS support.

[![CI](https://github.com/protectyr-labs/html2docx/actions/workflows/ci.yml/badge.svg)](https://github.com/protectyr-labs/html2docx/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)

Two-pass algorithm: collect element IDs for bookmarks, then convert
content to Word with formatting inheritance. Resolves CSS variables,
handles tables with headers, and generates cross-reference bookmarks.

## Quick Start

```bash
pip install git+https://github.com/protectyr-labs/html2docx.git
```

```python
from html2docx import HTMLToDocx

converter = HTMLToDocx()
converter.convert_file("report.html", "report.docx")
# => report.docx with headings, tables, formatting preserved
```

## Supported Elements

| Element | Support |
|---------|---------|
| Headings (h1-h6) | Full — with bookmark generation |
| Paragraphs | Full — with color and alignment |
| Tables | Full — thead/tbody, header shading, borders |
| Lists (ul/ol) | Full — nested, ordered and unordered |
| Inline formatting | Bold, italic, underline, color, code |
| Code blocks | Monospace font, preserved whitespace |
| Images | Embedded from local path |
| Links | Internal (bookmarks) + external |
| CSS variables | `var(--primary)` resolved from `:root` |

## Why This?

- **2-pass algorithm** — first collects all IDs, then converts with cross-reference bookmarks
- **CSS variable resolution** — handles `var(--color, fallback)` from `:root` declarations
- **Formatting inheritance** — RunFormat tracks bold/italic/color through nested inline elements
- **Color normalization** — dark greys (#1E293B) that look fine in browsers become unreadable in Word; auto-mapped to black

## Limitations

- No flexbox/grid layout (Word doesn't support CSS layout)
- No external CSS file loading (inline `<style>` blocks only)
- No JavaScript rendering (static HTML only)
- Images must be local files (no remote URL fetching)
- Complex nested tables may lose formatting

See [ARCHITECTURE.md](./ARCHITECTURE.md) for design decisions.

## License

MIT
