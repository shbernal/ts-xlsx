# Apply a style to a whole cell range in one call

Cluster: styles

## Scenario

A user wants to apply a fill (or any style facet) uniformly to a rectangular block of cells — shading
a header band, highlighting a region — **without iterating cell-by-cell**. Today the only path is to
loop over every row and column in the range and assign the style to each cell individually: verbose,
and easy to get wrong (off-by-one on the bounds, forgetting to materialize empty cells so an unwritten
cell in the block silently stays unstyled).

> Spec note, not a corpus case: this is an ergonomic API request with no failing behavior to assert —
> the per-cell loop already works. The durable value is the desired convenience surface and its design
> questions; the eventual behavior is assertable through the existing per-cell fill/style checks once
> the API exists, so no new adapter capability is needed now.

## Desired behavior

- A **range-styling call** that takes a range reference (`"B2:D5"`, or a `{tl, br}` pair) and a style
  payload and applies it to every cell in that block in one operation — including cells that were
  previously empty (they are materialized as styled-but-valueless, not skipped).
- **Merge vs. replace is explicit.** The default composes the payload onto each cell's existing style
  (set a fill without wiping fonts/borders/numFmt already there); a replace mode overwrites the whole
  style. The caller chooses; the library does not guess.
- **Bounded and predictable.** Styling a large range does not silently balloon the style table — it
  benefits from the same style deduplication every write path uses, so a uniform block resolves to one
  shared style index (see the dedup case).
- **Composable with existing single-cell styling** — a subsequent per-cell edit inside the range
  overrides just that cell, and range styling never corrupts merged-cell masters within the block.

## Open questions

- Surface shape: `worksheet.getRange("B2:D5").style = {...}`, a `worksheet.setRangeStyle(range, style,
  {mode})`, or both? Does a `Range` object become a first-class model type (enabling range-scoped
  values/formulas later) or is this a one-shot helper?
- Merge-vs-replace default: compose (least-surprise for "add a fill") vs. replace (matches assigning
  `cell.style`). This note leans compose-by-default.
- Empty-cell materialization: create styled empty cells across the whole block eagerly, or lazily via
  a column/row style so a 10,000-row block does not instantiate 10,000 cell objects?
- Interaction with column/row-level styles and with merged ranges that partially overlap the target.

Related: `shared-styles-deduplicated-in-written-package`, `per-cell-fill-isolation`,
`per-cell-font-isolation`, `column-level-value-type`, `worksheet-get-columns-range-accessor`,
`declarative-nested-column-headers`.
