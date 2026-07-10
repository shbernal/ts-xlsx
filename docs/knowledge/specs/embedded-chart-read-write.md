# Read and write embedded charts

Cluster: charts

## Scenario

Users want to read and write embedded charts (bar, column, line, pie, doughnut, scatter, area,
radar, …) inside a workbook. A chart lives as a DrawingML chart part anchored to a worksheet
drawing, referencing worksheet cell ranges as its data series and category axis. Consumers want to
open an application-produced workbook and enumerate/inspect its charts, and to declare new charts
that a spreadsheet application opens without repair and renders from live cell data. (Preserving an
unmodeled chart part verbatim across a round-trip is already captured by
`chart-parts-survive-template-roundtrip`; this note is the larger ask — a first-class, typed chart
model.)

> Spec note, not a corpus case: this is a substantial new feature with no failing behavior to assert
> yet. The durable value is the read/write chart model and its open design questions.

## Desired behavior

- **Read path**: parse the chart parts (DrawingML chart XML under the drawing relationships) and
  surface a typed chart model per worksheet — chart kind, title, plot area, one or more data series
  with their value/category/name range references, axis configuration, and the drawing anchor
  (cell-anchored position/size) tying the chart to the sheet.
- **Write path**: declare a chart on a worksheet by kind, series (each pointing at a cell range for
  values, optionally categories and a series-name cell), title, and anchor. On write, emit a valid
  chart part, the drawing part, and all relationships so the file opens without repair and renders
  from live worksheet data.
- **Round-trip**: a chart present in a loaded workbook survives a read→write unchanged (both the
  passthrough case and, once modeled, the typed case), with series range references intact.

## Open questions

- How much of the chart type space is first-class vs passthrough-only initially? (Bar/line/pie/
  scatter cover most demand.)
- Series data: only cell-range references (live data), or also inline/cached literal values?
- Is the chart model immutable-by-default with explicit mutation entry points (matching the fork's
  stance), and how does it compose with the drawing/anchor model shared with images?
- Combo charts, secondary axes, and stock/radar/surface types — in scope or deferred?

Related: `chart-parts-survive-template-roundtrip`, `pivot-table-parts-survive-roundtrip`,
`image-anchor-emu-from-real-column-geometry`.
