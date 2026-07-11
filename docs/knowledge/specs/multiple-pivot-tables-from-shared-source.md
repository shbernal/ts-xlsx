# A workbook must support multiple pivot tables (including from one shared source)

Cluster: pivot-tables

## Scenario

A user has one source data worksheet and wants several pivot views of it in the same workbook — one
summarizing sales by region, another the same rows by product, each with a different rows/columns/
values layout. The pivot feature must not cap a workbook at a single pivot table: adding a second (or
Nth) pivot table — over the same source range as an existing one, or a different source — must
succeed and round-trip to a package the application opens without repair, every table intact.

> Spec note, not a corpus case: ts-xlsx has no first-class pivot-table authoring API or adapter
> capability yet, so this is a forward-looking design target, not an immediately-assertable
> regression. When pivot support lands, the five constraints below are the concrete,
> implementation-blind properties to assert.

## Desired behavior

Adding N pivot tables must produce a valid package. The five correctness constraints — each of which
caused corruption or a repair prompt when a naive single-table implementation was generalized to N:

1. **Unique cache identity per pivot cache.** A cacheId shared across tables makes the application
   treat them as the same cache; each cache needs its own id.
2. **A distinct identity (UUID) per pivot table.** A reused pivot-table GUID collapses multiple
   tables into one; generate a fresh unique id per table.
3. **Complete cache field/sharedItems set.** A cache built only from the fields one table happens to
   use is insufficient when another table over the same source uses other fields; the cache should
   cover all source fields so any field can appear in any layout backed by it.
4. **Correct, unique per-table worksheet relationships.** Relationship targets must not be hardcoded
   to a single pivotTable part filename; each pivot sheet's rels point at its own pivotTable part with
   unique rel-ids (verifiable via worksheet-rels / package-part inspection).
5. **Non-mutating cache-field extraction.** Building the cache fields must not mutate the source
   worksheet's column data — a destructive splice-style construction corrupted source columns when a
   second table was created; repeated pivot creation over the same source must stay correct.

## Open questions

- **Cache sharing model:** when multiple tables target the identical source range, share one
  pivotCache (Excel's usual behavior — smaller files, coupled refresh) or give each a private cache?
  A cleaner design may dedupe caches by source range rather than always one-per-table.
- **Public API shape:** design the surface fresh (a workbook-level pivot builder, or `addPivotTable`
  with `sourceSheet`/`rows`/`columns`/`values`) rather than inheriting the experimental legacy shape.
- **Validation over prohibition:** instead of forbidding N>1, validate that source ranges, field
  names, and layouts are coherent and surface actionable errors.

Related: `pivot-table-round-trip-preservation`, `pivot-table-multiple-value-fields`,
`pivot-table-page-fields`, `pivot-table-aggregation-metrics`,
`pivot-table-preserve-worksheet-column-widths`, `pivot-cache-escapes-xml-special-characters`.
