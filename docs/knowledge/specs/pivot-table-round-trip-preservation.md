# Losing formatting/pivot table from loaded file

## Pivot table preservation and support

### Problem
When a workbook containing pivot tables is loaded and then written back out, the pivot tables are destroyed. The library historically had no data model for pivot-table parts, so on write those parts are dropped: the pivot region degrades into a plain sheet of static cell values and pivot-associated formatting (notably theme-driven font colors such as white header text) is lost. This is one of the most-requested and longest-running gaps reported against upstream, reproduced repeatedly across many library versions and still present at the last upstream release.

### Desired behavior (tiered)

**Tier 1 — lossless pass-through (highest value, lowest cost).**
Loading a workbook that contains pivot tables and writing it back — with no edits, or with edits confined to unrelated worksheets — must preserve the pivot tables such that Excel still opens them as live, interactive pivots. Concretely, the OOXML parts that constitute a pivot must survive the round-trip intact and correctly re-referenced:
- `xl/pivotTables/pivotTable*.xml` (pivot table definitions)
- `xl/pivotCache/pivotCacheDefinition*.xml` and `pivotCacheRecords*.xml` (the cached source data — the hard part, per prior art)
- the `pivotCacheDefinition` relationships in the workbook, the `pivotTables` relationships in each hosting worksheet, and the corresponding `[Content_Types].xml` overrides
Even a workbook the library does not otherwise understand should not have these parts silently deleted; unknown-but-referenced parts should be carried through rather than dropped.

**Tier 2 — read access.**
Expose parsed pivot-table definitions (source range / cache, row/column/data/filter field placement, and the cached field items) so callers can inspect existing pivots.

**Tier 3 — authoring.**
Allow creating new pivot tables programmatically. The acknowledged hard part is not the pivot *definition* XML but generating the *cached* content (`pivotCacheRecords`) that Excel stores alongside — a from-scratch author must materialize the cache from the source range.

### Related secondary defect
Independently reported alongside this: font color flipping (white → black) on round-trip. Upstream attributed part of that to a theme-handling bug that was claimed fixed separately; treat theme color fidelity on round-trip as its own concern rather than folding it into pivot support.

### Prior art / open questions
- Prior notes acknowledge that recreating the pivot cache records is the central difficulty for authoring; pass-through avoids that entirely and should be delivered first.
- Open question: scope of Tier 1 — is byte-faithful copy of the pivot parts acceptable, or must the library re-resolve/renumber relationships when worksheets are added/removed around them?
- Open question: how to represent unmodeled-but-preserved parts in the object model so they survive edits without the library needing to fully understand them.
