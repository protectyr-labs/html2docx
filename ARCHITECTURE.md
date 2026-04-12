# Architecture

## Why 2-Pass Conversion

The converter uses a two-pass algorithm:

- **Pass 1** walks the entire HTML tree and collects all element `id` attributes into a set.
- **Pass 2** converts HTML elements to Word objects, creating bookmarks for any element whose `id` was collected in Pass 1.

This is necessary because Word bookmarks require a numeric ID at creation time, and internal links (`<a href="#section">`) need to know whether the target exists before the target element has been processed. By collecting all IDs first, Pass 2 can create bookmarks and cross-references in a single forward pass without backpatching.

## Why RunFormat Dataclass

HTML inline formatting nests arbitrarily: `<em><strong><span style="color:red">text</span></strong></em>`. Each nesting level can add formatting but should never remove formatting applied by a parent.

`RunFormat` is a simple dataclass that accumulates formatting state. When entering a nested element, the current `RunFormat` is copied and the new element's formatting is applied on top. This copy-on-descent pattern means:

- Bold inside italic produces bold+italic (not just bold)
- Color set by a parent persists into children unless overridden
- No formatting is lost when exiting a nested element

## Why CSS Variable Resolver

Modern HTML generators (report templates, static site generators, design systems) use CSS custom properties extensively:

```css
:root { --primary: #1a2744; --danger: rgb(220, 53, 69); }
```

Word documents have no concept of CSS variables. The `CSSVariableResolver` parses `:root` declarations and replaces `var(--name)` references with their resolved values before any color parsing occurs. This means `style="color: var(--primary)"` correctly produces a Word run with `RGBColor(0x1a, 0x27, 0x44)`.

The resolver supports fallback syntax: `var(--missing, #333)` returns `#333` when `--missing` is undefined.

## Element Dispatch

`_process_children` is the central dispatcher. It walks a parent element's children and routes each to the appropriate handler:

| Tag | Handler | Notes |
|-----|---------|-------|
| h1-h6 | `_process_heading` | Heading style + font size + optional bookmark |
| p | `_process_paragraph` | Inline children + alignment |
| table | `_process_table` | thead/tbody detection, cell shading |
| ul, ol | `_process_list` | Recursive for nesting, bullet vs number style |
| pre | `_process_code_block` | Monospace font, gray shading |
| blockquote | `_process_blockquote` | Left indent + italic |
| hr | `_process_hr` | Bottom border paragraph |
| img | `_process_image` | Local files only |
| div, section, etc. | Recurse into `_process_children` | Transparent containers |
| inline tags | `_process_inline` via paragraph wrapper | Block-level inline fallback |

## Known Limitations

- **No float/position layout.** Word paragraphs are strictly sequential. CSS `float`, `position: absolute`, and similar layout properties are ignored.
- **No flexbox/grid.** CSS Grid and Flexbox have no Word equivalent. Content is linearized.
- **No external CSS files.** Only `<style>` blocks within the HTML are parsed. External `.css` files referenced via `<link>` are not fetched.
- **No remote images.** Only local file paths in `<img src>` are embedded. URLs are replaced with alt text.
- **No media queries.** Responsive breakpoints are irrelevant for Word output.
- **No JavaScript.** Dynamic content must be pre-rendered before conversion.
- **Simple table model.** No colspan/rowspan support. Complex merged-cell layouts will not render correctly.
