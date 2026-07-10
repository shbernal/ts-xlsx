# Sort worksheet rows by one or more key columns

Cluster: core-model

## Scenario

A user has a worksheet of tabular data and wants to reorder its data rows so a chosen column sorts
ascending or descending, optionally over a sub-range that keeps a header row fixed. The library
exposes no first-class sort, so users hand-roll it: read every cell into an array-of-arrays, sort by
a comparator on the chosen column, then remove and re-insert rows via row-splices. This workaround
is lossy and fragile — it flattens rows to bare values, so per-cell/per-row styles, heights, notes,
merged regions, and formulas do not travel with their row (formulas in particular move without their
references being reconsidered).

> Spec note, not a corpus case: this is a new capability with no failing behavior to assert yet.
> The durable value is the desired operation and the many semantic decisions it forces.

## Desired behavior

- A first-class sort reorders **whole rows as units**, so style, height, notes, and merged regions
  travel with their row rather than being flattened to values.
- Sort by one or more key columns, each with its own ascending/descending direction, over an
  optional row range that preserves a header.
- **Stable**: equal keys keep their original relative order.
- The comparator defines a **total, predictable order across mixed value types** — numbers, strings,
  booleans, dates, rich text, hyperlink/formula result values, and empty/null cells — with defined
  placement for empties and defined case/locale handling for strings.
- Formulas in moved rows have a **documented policy**: either relative references are rewritten to
  follow the row's new position, or the operation is explicitly value-preserving only.

## Prior art

Spreadsheet applications offer range sort with header detection, multi-level keys, custom orders,
and case-sensitivity toggles. OOXML also persists a `sortState` on tables/autofilters describing the
last-applied sort — a complete solution could optionally emit that so the applied order is recorded
in the file, distinct from a purely in-memory reorder.

## Open questions

- Is sorting an eager in-memory row reorder, a persisted `sortState` on a table/autofilter, or both?
- Default placement of empty/null cells (first vs last), and is it configurable?
- How are rows spanning merged cells that cross the sort range handled?
- Does the API key on column index, column letter, or a table column name?
- Formula-reference rewriting vs a value-only guarantee.

Related: `splice-rows-carries-styles-on-shifted-rows`, `splice-rows-preserves-merged-cells`,
`row-iteration-early-termination`.
