# html2docx

HTML to Word document converter with CSS variable resolution, bookmarks, and formatting inheritance.

## Why This Exists

Reporting pipelines frequently generate HTML output that needs to be delivered as Word documents. Existing conversion tools either lose formatting, ignore CSS variables, or fail on nested inline elements. This library handles the common cases correctly using a 2-pass algorithm that first collects element IDs for bookmark generation, then converts content with full formatting inheritance.

## Quick Start

```bash
pip install html2docx-converter
```

```python
from html2docx import HTMLToDocx

converter = HTMLToDocx()

# From string
converter.convert("<h1>Report</h1><p>Content here.</p>", "report.docx")

# From file
converter.convert_file("input.html", "output.docx")
```

## Supported Elements

| HTML Element | Word Output | Notes |
|---|---|---|
| `<h1>` - `<h6>` | Heading 1-6 | Font size scaled per level |
| `<p>` | Paragraph | Supports text-align |
| `<strong>`, `<b>` | Bold run | Nests correctly |
| `<em>`, `<i>` | Italic run | Nests correctly |
| `<u>` | Underlined run | |
| `<code>` | Consolas 9pt run | Inline code |
| `<pre><code>` | Consolas block | Gray background shading |
| `<table>` | Word table | thead/tbody, header shading |
| `<ul>`, `<ol>` | Bullet/Number list | Nested lists supported |
| `<blockquote>` | Indented italic | |
| `<hr>` | Bottom border line | |
| `<img>` | Inline image | Local files only |
| `<a>` | Plain text | Bookmark targets via `id` |
| `<span>` | Styled run | Inherits parent formatting |
| `<mark>` | Yellow highlight | |

## CSS Variable Support

The converter resolves CSS custom properties defined in `:root`:

```html
<style>
:root {
  --primary: #1a2744;
  --accent: rgb(255, 102, 0);
}
</style>
<h1 style="color: var(--primary)">Styled Heading</h1>
<p style="color: var(--accent)">Accented text</p>
```

Variables are resolved before color parsing. Supports `#hex` (3 and 6 digit) and `rgb()` formats. Fallback values work: `var(--missing, #333)`.

## API Reference

### `HTMLToDocx`

Main converter class.

```python
converter = HTMLToDocx()
```

#### `convert(html: str, output_path: str) -> str`

Convert an HTML string to a DOCX file. Returns the absolute output path.

#### `convert_file(html_path: str, output_path: str) -> str`

Convert an HTML file to a DOCX file. Returns the absolute output path.

#### `stats: dict`

After conversion, contains counts: `{"headings": N, "paragraphs": N, "tables": N, "lists": N}`.

### `CSSVariableResolver`

Standalone CSS variable resolver.

```python
from html2docx import CSSVariableResolver

resolver = CSSVariableResolver(":root { --brand: #336699; }")
resolver.resolve("var(--brand)")        # "#336699"
resolver.resolve("var(--x, fallback)")  # "fallback"

CSSVariableResolver.parse_color("#abc")           # "AABBCC"
CSSVariableResolver.parse_color("rgb(255, 0, 0)") # "FF0000"
```

### `RunFormat`

Dataclass tracking inline formatting state through recursive processing.

```python
from html2docx import RunFormat

fmt = RunFormat(bold=True, italic=True, color="FF0000")
child_fmt = fmt.copy()
child_fmt.underline = True  # Original unchanged
```

## Development

```bash
git clone https://github.com/protectyr-labs/html2docx.git
cd html2docx
pip install -e ".[dev]"
pytest
```

## License

MIT
