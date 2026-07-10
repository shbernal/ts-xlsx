# AutoFilter should support suppressing the dropdown button on chosen columns

Cluster: tables

## Scenario

A worksheet has an AutoFilter over a horizontal range of header cells. The author wants the filter
dropdown button on some columns but not others — a common need when one logical field spans several
columns or uses a merged header, so that only one dropdown (or none) shows for that spanned region.
OOXML expresses this with per-column `filterColumn` entries carrying `hiddenButton="1"`: the column is
part of the filter range but shows no control. The current writer emits only the `autoFilter` range
(`<autoFilter ref="A1:C1"/>`) and offers no way to hide individual buttons, so every column in the
range gets a dropdown.

> Spec note, not a corpus case: this is an additive capability that does not exist yet — there is no
> current behavior to assert a baseline against. The durable value is the desired contract and its
> design questions; it becomes a corpus case (write → inspect per-column button state → round-trip)
> once the authoring surface lands, and the adapter's `autoFilter` inspection fact is extended to read
> per-column `hiddenButton` state at that point.

## Desired behavior

- **An AutoFilter accepts an opt-in set of columns whose dropdown button is hidden.** Writing such a
  worksheet emits a `filterColumn` with `hiddenButton="1"` for precisely those columns and leaves the
  rest with their default button.
- **Hidden-button columns are addressed relative to the AutoFilter's left edge**, matching the OOXML
  `filterColumn colId` (zero-based from the range start), so the intended columns are affected
  regardless of where the range begins on the sheet.
- **The default is unchanged.** An AutoFilter with no hidden-button columns behaves exactly as today —
  a dropdown on every column in the range.
- **Round-trip fidelity.** Loading and re-writing a workbook preserves both the AutoFilter range and
  the set of columns whose buttons are hidden; the hidden-button state is not silently dropped.
- **Composes with active filter criteria.** Hiding a button is independent of whether that column also
  carries filter criteria — a column can have criteria applied while its button is hidden, and reading
  criteria (see the read-criteria case) must not be disturbed by the hidden-button flag.

## Open questions

- Authoring surface: an `autoFilter: { from, to, columns: [{ column, filterButton: false }] }` shape,
  or a separate per-column call? Column identification by address vs. by zero-based offset in the API
  (the file format uses the offset; the API should probably accept an address and translate).
- Interaction with merged header cells (the motivating case) — does hiding buttons on the covered
  columns of a merge become the ergonomic default, or stay fully explicit?
- Whether this pulls in the broader `filterColumn` surface (custom/dynamic/top-10 filters), or ships
  narrowly as just button visibility first.

Related: `autofilter-range-is-bounded-rectangle`, `read-worksheet-with-autofilter-criteria`,
`autofilter-emits-filter-database-defined-name`, `multi-table-autofilter-survives-roundtrip`.
