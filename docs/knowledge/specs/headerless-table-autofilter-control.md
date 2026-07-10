# AutoFilter on a table is explicitly controllable, independent of the header row

Cluster: tables

## Scenario

A developer wants a worksheet table purely for its consistent styling — banded rows, a table
theme — with no filter dropdowns and, in some cases, no header row at all: a clean data block that
Excel opens without a repair prompt. Two related frustrations recur:

1. A **header-less** table (`headerRow: false`) still gets an AutoFilter emitted against the
   now-absent header, producing an internally inconsistent table part that Excel repairs on open.
   That specific corruption is already pinned by the corpus case
   `table-headerless-omits-autofilter` (a header-less table must set `headerRowCount=0` and emit no
   `autoFilter`).
2. Even a **header-bearing** table has its AutoFilter forced on implicitly — there is no way to say
   "give me a styled table with a header but no filter UI." The only escape is to abandon the table
   feature and hand-style cells.

The durable question this note captures is the missing *design surface*: filter presence should be
an explicit, first-class choice, not an implicit consequence of whether a table has a header.

> Spec note, not a corpus case: the corruption path (header-less + forced filter) is already locked
> as a corpus case. What remains is an API-shape decision — how a caller expresses "table without a
> filter" — which is design material, not a malformed-output bug.

## Desired behavior

- **Filter emission follows an explicit flag, not the header's presence.** A table can be written
  with a header and no AutoFilter, or with no header and no AutoFilter, or with a header and an
  AutoFilter — all three are expressible and all three produce valid OOXML Excel opens without
  repair.
- **`headerRow: false` conflates two concerns today** (header presence *and* filter presence);
  separate them so the two axes are independent. A header-less table never carries an AutoFilter
  (it is not legal there); a header-bearing table carries one only when the caller asks.
- **Round-trip preserves both axes**: a table written header-less stays header-less on reload, and a
  table written filter-less does not silently regain a filter.
- **Styling stays valid without a header**: banded-row / first-column style flags remain meaningful
  on a header-less table; the table style definition is still emitted.

## Open questions

- The public surface: a dedicated `filterButton`/`autoFilter` boolean on the table definition versus
  inferring from `headerRow`. A separate flag is clearer and lets "header + no filter" exist.
- Default behavior when the flag is unset: match Excel (a header-bearing table defaults to showing
  filter buttons) but never emit a filter for a header-less table. Document the chosen default.
- Interaction with a totals row: confirm totals-row handling stays consistent when the filter is
  absent.

Related: `table-column-width-in-definition`, `streaming-writer-table-support`,
`table-handle-direct-property-access`.
