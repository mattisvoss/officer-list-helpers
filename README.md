# officer-list-helpers

Helper functions that add bullet and numbered list support to the R
[officer](https://github.com/davidgohel/officer) package. Works with any
template and any custom paragraph style.

## The problem

Officer lets you create Word documents from R, but adding bullet or numbered
lists requires a pre-defined list style in your Word template. If you have a
custom paragraph style (say "Report Body") and want some paragraphs in that
style to also be bulleted, you're stuck — there's no way to layer list
formatting on top of an arbitrary style.

These helpers solve that by injecting the list XML directly into the paragraphs
officer creates, so your custom styling is preserved.

## Quick start

```r
source("officer_list_helpers.R")

doc <- read_docx("my_template.docx")

# Bullet list using your custom style
doc <- list_add_par(doc, "First point",  style = "Report Body", list_type = "bullet")
doc <- list_add_par(doc, "Sub-point",    style = "Report Body", list_type = "bullet", ilvl = 1L)
doc <- list_add_par(doc, "Second point", style = "Report Body", list_type = "bullet")

# Normal paragraph (not a list item) — just use officer as usual
doc <- body_add_par(doc, "Some normal text.", style = "Report Body")

# End the bullet list, start a numbered list
doc <- list_end(doc)
doc <- list_add_par(doc, "Step one", style = "Report Body", list_type = "decimal")
doc <- list_add_par(doc, "Step two", style = "Report Body", list_type = "decimal")

print(doc, target = "output.docx")
```

## Installation

No installation needed. Copy `officer_list_helpers.R` into your project and
source it:

```r
source("officer_list_helpers.R")
```

### Dependencies

- [officer](https://cran.r-project.org/package=officer) — the Word document
  builder these helpers extend
- [xml2](https://cran.r-project.org/package=xml2) — for reading and modifying
  the XML inside .docx files

Install them with:

```r
install.packages(c("officer", "xml2"))
```

## API reference

### `list_add_par(x, value, style, list_type, ilvl, pos)`

Add a plain text paragraph as a list item. The most common entry point.

```r
doc <- list_add_par(doc, "Buy groceries", style = "Normal", list_type = "bullet")
doc <- list_add_par(doc, "Milk",          style = "Normal", list_type = "bullet", ilvl = 1L)
```

| Parameter   | Type      | Default    | Description                                      |
|-------------|-----------|------------|--------------------------------------------------|
| `x`         | rdocx     | (required) | Document object from `read_docx()`               |
| `value`     | character | (required) | The paragraph text                               |
| `style`     | character | `NULL`     | Paragraph style from your template               |
| `list_type` | character | `"bullet"` | `"bullet"` or `"decimal"`                        |
| `ilvl`      | integer   | `0L`       | Indent level: 0 = top, 1 = sub-item, 2 = sub-sub |
| `pos`       | character | `"after"`  | `"after"`, `"before"`, or `"on"`                 |

### `list_add_fpar(x, value, style, list_type, ilvl, pos)`

Add a formatted paragraph (mixed bold, italic, etc.) as a list item. Build
your content with `fpar()` and `ftext()` from officer.

```r
formatted <- fpar(
  ftext("Important: ", prop = fp_text(bold = TRUE)),
  ftext("buy milk")
)
doc <- list_add_fpar(doc, formatted, list_type = "bullet")
```

### `list_add_blocks(x, blocks, list_type, ilvl, pos)`

Add multiple formatted paragraphs as list items in one call. Each paragraph
in the `block_list` becomes a separate list item.

```r
items <- block_list(
  fpar(ftext("First point")),
  fpar(ftext("Second point")),
  fpar(ftext("Third point"))
)
doc <- list_add_blocks(doc, items, list_type = "decimal")
```

### `list_end(x)`

End the current list. The next `list_add_*` call will start a new list that
counts from 1.

You only need this between two lists **of the same type**. Switching from
bullet to decimal (or vice versa) automatically restarts.

```r
doc <- list_add_par(doc, "A", list_type = "decimal")  # renders as 1.
doc <- list_add_par(doc, "B", list_type = "decimal")  # renders as 2.
doc <- list_end(doc)
doc <- list_add_par(doc, "C", list_type = "decimal")  # renders as 1. (restarted)
```

### `list_inspect(x)`

Print a human-readable summary of every numbering definition in the document.
Useful for debugging and for understanding what your template provides.

```r
list_inspect(doc)
```

Output looks like:

```
=== Format templates (abstractNum) ===
  abstractNumId = 0
    level 0: format = bullet      display = •
    level 1: format = bullet      display = ◦

=== List instances (num) ===
  numId = 1  ->  abstractNumId = 0
  numId = 2  ->  abstractNumId = 0  [restarts at 1]
```

## How it works

If you're curious about the internals, the source file contains a detailed
tutorial covering:

1. **How .docx files work** — they're zip archives of XML files
2. **How Word lists work** — the two-layer numbering system (abstractNum
   format templates and num instances)
3. **How numbering restart works** — the `<w:lvlOverride>/<w:startOverride>`
   mechanism
4. **How XPath works** — the query language for finding things in XML
5. **How officer fits in** — unpacking the .docx, cursor-based editing, and
   our strategy of layering list XML on top of officer's output

Read the comments in `officer_list_helpers.R` — they're designed to teach you
everything you need to know, even if you've never worked with XML before.

## Running the tests

```bash
Rscript test_officer_list_helpers.R
```

The test suite covers bullet creation, numbered lists, indent levels, restart
behaviour, style preservation, formatted paragraphs, block lists, round-trip
save/reload, and edge cases. Each test is annotated with comments explaining
what it checks and why.

## License

MIT — see [LICENSE](LICENSE).
