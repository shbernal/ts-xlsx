# A pivot table can preserve the worksheet's column widths instead of its style's sizing

Cluster: pivot-tables

## Scenario

A user builds a spreadsheet with a pivot table and has carefully set custom column widths on the
worksheet (a wide label column, narrow data columns, for a print or dashboard layout). By default the
application applies the pivot table style's own width/height formatting, which overrides the
worksheet column widths, so the custom sizing is lost on open. The user wants to tell the pivot table
to preserve the worksheet's column widths instead of letting the pivot style reset them.

> Spec note, not a corpus case: pivot-table authoring is not part of the ts-xlsx declarative spec or
> adapter surface yet, so this is a sub-feature of the broader "generate pivot tables" capability
> that must be designed first (see `multiple-pivot-tables-from-shared-source`). Recorded as the
> target for when pivot authoring lands.

## Desired behavior

- **The author chooses whether the pivot style controls sizing.** When emitting a pivot table, let
  the author decide whether the pivot table style controls column widths/row heights, or whether the
  worksheet's own widths/heights are preserved.
- **This maps to OOXML `pivotTableDefinition/@applyWidthHeightFormats`** (Part 1 §18.10): `1` (Excel
  default) applies the pivot style's width/height formats; `0` preserves the worksheet's. The library
  default should match Excel's (apply the style formats) so existing output is unchanged, with an
  opt-in to disable it.
- **Round-trip assertion:** with the option off, worksheet column widths must survive the write; with
  it on (default), the attribute is present/at its default.

## Open questions

- **Typed API over the raw attribute:** expose a real boolean (`true`/`false`) mapped to the XML
  attribute internally, not the OOXML string `'0'`/`'1'`.
- **Author-facing name:** `applyWidthHeightFormats` is faithful to OOXML but opaque; a clearer name
  (`preserveColumnWidths` / `applyPivotStyleSizing`) may serve users better, with the OOXML attribute
  as an implementation detail.
- **Dependency on the broader pivot API:** this option only exists once pivot authoring itself is
  designed; it should be a per-pivot-table option alongside `sourceSheet`/`rows`/`columns`/`values`.

Related: `multiple-pivot-tables-from-shared-source`, `pivot-table-round-trip-preservation`,
`column-width-round-trips-exactly`, `default-font-must-not-be-assumed-for-column-widths`.
