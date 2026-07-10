# Copy rows/ranges across worksheets preserving styles, formulas, and merges

Cluster: core-model

## Scenario

A user consolidates several files into one workbook by copying rows from source worksheets into a
destination, expecting the copied rows to carry over not just values but cell styles, number formats,
formulas, and merged-cell ranges. Naive value-only copying loses formatting; and copying rows that
contain shared-formula member cells whose master lies outside the copied region trips the same
"master must exist above/left of clone" failure locked as a known-open in
`shared-formula-master-survives-roundtrip-and-splice`. Users resort to `Object.assign` over row
objects, which happens to drag style internals along but silently breaks on formulas and merges.

> Spec note, not a corpus case: this is a first-class API that does not exist yet. The durable value
> is the copy semantics — what must be preserved and how the hard edges (shared formulas, merges,
> cross-workbook style tables) are handled.

## Desired behavior

- A well-typed way to copy a row (or a range of rows/cells) from one worksheet to another — the same
  sheet, another sheet in the same workbook, or a sheet in a different loaded workbook — preserving
  values, cell styles, number formats, alignment/fill/borders, formulas, and merged ranges.
- **Resilient to shared formulas**: when a copied cell participates in a shared formula whose master
  lies outside the copied region, the copy must not crash — it either materializes the concrete
  formula for that cell (expanding the shared formula into a plain one) or re-anchors a new
  shared-formula master within the destination. It never throws the master-position error.
- **Merged ranges** intersecting the copied region are reconstructed in the destination at the
  corresponding offset.
- **Cross-workbook** copies remap style/number-format references into the destination's style table
  rather than dangling indices into the source's.

## Open questions

- API shape: `dst.addRow(srcRow)` that deep-copies, a `copyRange(srcRange, dstAnchor)` helper, or a
  `Range` model (shared with `sort-rows-by-column` / `set-style-over-cell-range`)?
- Formula policy default: expand shared formulas to concrete per-cell formulas (simple, lossless in
  value) vs re-anchor a shared master (compact) — and are relative references rebased to the new
  position?
- Does a cross-workbook copy also carry defined names / conditional formats referenced by the copied
  rows, or only the cell-local state?

Related: `shared-formula-master-survives-roundtrip-and-splice`, `sort-rows-by-column`,
`set-style-over-cell-range`, `worksheet-model-preserves-merged-cells`, `copyWorksheetModel`.
