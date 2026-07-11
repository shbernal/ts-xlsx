# PivotTable page fields (report filters)

Cluster: pivot

## Scenario

A user builds a pivot table and wants to filter the entire table by a field's value — the pivot
"page field" / report filter that renders above the pivot body as a dropdown. For example a `latest`
field with values 0/1 could default to showing only `latest = 1` rows, with the option to select
`latest = 0` or all values from the dropdown. Today the pivot supports row, column, and value fields,
but not page fields, so this whole-table filter cannot be authored.

> Spec note, not a corpus case: page fields do not exist in the pivot model yet, so there is no
> behavior to assert. The durable value is the page-field model and its OOXML mapping. It becomes a
> corpus case once the authoring API accepts page fields and the written pivot definition's
> `<pageFields>` block (and any default selected item) is asserted.

## Desired behavior

- **A pivot accepts one or more page fields** in addition to rows/columns/values — fields that filter
  the entire pivot rather than grouping it. Each named page field maps to a source field.
- **A default selected item is expressible per page field** (e.g. default the `latest` filter to `1`),
  and is serialized so the host opens the pivot pre-filtered to that item; omitting a default leaves
  the filter on "(All)".
- **Page fields serialize into the pivot definition** as a `<pageFields>` block whose entries
  reference the pivot-cache field and carry the selected item index, and the field is marked as a page
  axis (`axis="axisPage"`) in the pivot fields — so the host renders the report-filter dropdown above
  the pivot.
- **The existing row/column/value behavior is unchanged** when no page fields are supplied.

## Prior art (OOXML)

In the pivot table definition, a field placed on the report-filter axis is marked `axis="axisPage"`
in its `<pivotField>`, and the `<pageFields count="…">` collection carries one `<pageField fld="…"
item="…"/>` per filter, where `item` is the index (within the cache field's shared items) of the
currently-selected value; its absence means "(All)". The referenced values must exist in the
pivot-cache field's shared items.

## Open questions

- Public API shape: a `pages: ['latest']` array plus an optional `pageDefaults: { latest: 1 }` map,
  versus a structured per-page object. How an unknown default value (not present in the source column)
  is handled — reject at author time vs. fall back to "(All)".
- How the pivot cache's shared items are populated so a default `item` index resolves to the intended
  value, and whether multi-select page filters (more than one selected item) are in scope.
- Interaction with the `metric` / aggregation choice (`pivot-table-aggregation-metrics`) and multiple
  value fields (`pivot-table-multiple-value-fields`) when a page filter is also present.

Related: `pivot-table-multiple-value-fields`, `pivot-table-aggregation-metrics`,
`pivot-table-round-trip-preservation`, `pivot-and-slicer-parts-survive-roundtrip`.
