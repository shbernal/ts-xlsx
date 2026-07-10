# PivotTable support for multiple value fields and the "Values" axis

Cluster: pivot

## Scenario

A user builds a pivot table from a tabular source and wants to aggregate several numeric measures at
once — e.g. from Initial Cost, VAT, and Total, produce a summed value per grouping row. When multiple
value fields are requested and no explicit column-grouping field is given, the pivot should place the
measures side by side along a synthesized "Values" axis, the way spreadsheet applications do. Today a
pivot definition handles the single-measure shape but not multiple measures or the optional-columns
case.

> Spec note, not a corpus case: this extends an unbuilt/partial pivot feature. The durable value is
> the multi-measure model and the "Values" placeholder axis semantics.

## Desired behavior

- A pivot definition **accepts more than one value/measure field**. Each supplied measure is emitted
  as its own data field with the chosen aggregation (e.g. sum).
- When **two or more measures** are present, the generated pivot places them along a synthesized
  **"Values" axis** — reproducing the native placeholder a spreadsheet app inserts into the column
  area so the measures render side by side.
- The **column-grouping axis is optional**: supplying no column fields is a valid pivot; the writer
  defaults to the "Values" placeholder axis rather than erroring.
- The **single-measure-with-column-grouping** baseline (one value field distributed across a real
  column field's distinct values) continues to work unchanged.

## Open questions

- Per-measure aggregation functions (sum/count/avg/min/max) and whether they are individually
  configurable, or one function for all measures initially.
- Where the "Values" axis sits when explicit column fields ARE present (measures nested under the
  column grouping vs alongside).
- Number-format inheritance: does each measure carry the source column's number format into the
  pivot's value cells?
- How much of this is authored vs round-tripped from an application-created pivot (the existing pivot
  work is passthrough-oriented — see `pivot-table-parts-survive-roundtrip`).

Related: `pivot-table-parts-survive-roundtrip`, `pivot-and-slicer-parts-survive-roundtrip`,
`embedded-chart-read-write`.
