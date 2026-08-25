# Export a worksheet (or range) to an HTML fragment

Cluster: xlsx-io

## Scenario

A developer has read or built a workbook in memory and wants to display it in a web page without
hand-writing DOM or table markup: a `<table>` of cell values that honors merged cells and, ideally,
basic fonts, fills, borders, column widths, and number formatting. This is the inverse of authoring a
cell from a markup fragment (`html-fragment-to-rich-text-cell-value`), where the worksheet is the
source and HTML is the output.

> Spec note, not a corpus case: this is an opt-in convenience feature with no failing behavior to
> assert yet. The durable value is the export's feature envelope and its formatting fidelity.

## Desired behavior

- An opt-in export turns a worksheet, or a cell range, into an HTML fragment: at minimum a `<table>`
  of computed cell values with `rowspan` and `colspan` for merged ranges.
- A richer tier maps style facts already modeled (font weight, italic and color; fill color; borders;
  horizontal and vertical alignment; column widths; row heights) to inline styles or classes, and
  applies **number-format-aware** value formatting so displayed text matches the spreadsheet.
- **Formulas render as their cached result**, not the formula text.
- Output is a **string, or streamable chunks, with no live DOM dependency**, so it works server-side.

## Prior art

Standalone converters and spreadsheet-viewer widgets already do value-plus-basic-style-to-table
conversion and define the reasonable envelope: merged cells, basic styles, number formats. This is
a display and export helper, not a layout engine.

## Open questions

- Feature tiers: values-only, values plus styles, or full-fidelity? What is the default, and how is
  the richer tier opted into?
- Style mapping: inline styles, giving a self-contained fragment, or classes plus a supplied
  stylesheet, giving smaller output the caller themes?
- Number-format fidelity: reuse the same format engine that renders cells for write, so HTML and the
  spreadsheet agree on displayed text.
- Scope boundary: images, charts, and conditional formatting in the HTML output, or values plus static
  styles only?

Related: `html-fragment-to-rich-text-cell-value`, `worksheet-model-preserves-merged-cells`,
`set-style-over-cell-range`.
