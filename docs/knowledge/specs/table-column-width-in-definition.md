# A table column definition should accept an inline display width

Cluster: tables

## Scenario

An author declares a table with its column definitions (name, totals-row function, filter button)
and wants to set each column's display width in the same place, rather than declaring the table and
then separately setting the underlying worksheet columns' widths by position. Today the width has to
be applied out-of-band on the worksheet column that the table happens to occupy, which is
error-prone (the author must track which sheet-column index each table column maps to) and splits one
logical declaration across two APIs.

> Spec note, not a corpus case: the two-step path already works and produces correct output, so there
> is no incorrect behavior to assert — this is an ergonomics/API-shape proposal. The durable value is
> the equivalence contract and where width belongs in the table authoring surface.

## Desired behavior

- **A table column definition accepts an optional `width`.** Declaring it inline is exactly
  equivalent to setting the corresponding worksheet column's width by position after the table is
  created: same rendered column widths, same OOXML output (`<col>` width on the sheet; the table part
  is unaffected — width is not a table-column property in the format). The author may choose the
  ergonomic inline form with no behavioral surprise.
- **Width lives on the sheet column, not the table part.** OOXML tables (`<tableColumn>`) carry no
  width; column width is a worksheet `<col>` attribute. The inline table-column `width` is sugar that
  the writer resolves onto the sheet column the table spans — it does not invent a non-standard table
  attribute.
- **Composition is predictable.** An inline table-column width and an explicit worksheet-column width
  for the same column must not both silently apply and fight; define precedence (most naturally, an
  explicit later worksheet-column width wins, or the two are the same knob) and document it.

## Open questions

- Precedence when both an inline table-column width and a worksheet `columns[].width` target the same
  column — last-write-wins, table-definition-wins, or treat them as one property?
- Does the inline width extend to other per-column display properties (number format, style, hidden)
  that are likewise really worksheet-column concerns, or is width the only sugar worth adding?
- Read-back: when a table is read from a file, are its columns' widths surfaced on the table-column
  view (resolved from the sheet columns) or only on the worksheet columns?

Related: `worksheet-columns-mutable-array-ergonomics`, `table-handle-direct-property-access`,
`streaming-writer-table-support`, `column-definition-type-is-partial-on-write`.
