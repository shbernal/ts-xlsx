# Granting sort/AutoFilter on a protected sheet needs an unlocked window, not just the flag

Cluster: protection

## Scenario

An author protects a worksheet where most cells are locked, but wants end users to still sort and
filter the data range without unprotecting the sheet. They enable the `sort` and `autoFilter`
permissions when turning on protection and expect sorting to work. It doesn't: the application still
refuses to sort, because sorting rewrites cell values and the cells are locked. A sheet protected
*manually* in Excel behaves differently only because the user typically also left the sortable range
unlocked (or used Allow-Edit-Ranges). The reported "programmatic vs manual" gap is not a hidden
format quirk — it is which cells are left unlocked and which protection surface is emitted.

> Spec note, not a corpus case: the flag *encoding* is already correct and locked by
> `sheet-protection-permits-requested-operations` (granting sort emits `sort="0"`). What is missing
> is the higher-level semantics — the flag alone is insufficient, and the fix requires either
> per-cell unlock state or a `<protectedRanges>` editable window, neither of which the library models
> as an authoring surface. The exact attribute+lock combination the application honors must be
> verified empirically before any assertion; this records the design target.

## Desired behavior

- **A protected sheet with sort/AutoFilter granted actually permits those operations** in the
  application, matching a manually-protected sheet — not merely setting the permission attributes.
- **The mechanism is an unlocked window, not just a flag.** Sort/AutoFilter over a protected range
  requires that the reorderable data cells are themselves unlocked (`locked=false`), OR that a
  `<protectedRanges>` entry carves out an editable region within the otherwise-protected sheet.
  Setting `sort`/`autoFilter` on `<sheetProtection>` while every covered cell stays locked is not
  enough — the application still refuses to reorder locked cells.
- **The OOXML surface involved:** the `<sheetProtection>` attributes (`sort`, `autoFilter`,
  `formatCells`, the `sheet` flag), the per-cell/per-style `locked` protection property, and
  optionally `<protectedRanges>` definitions.

## Open questions

- Should the library expose a high-level convenience — a `sortableRange` / `allowEdit` option on
  protection — that automatically unlocks the relevant cells (or emits a protected range) so the
  common case "protect everything but let users sort this table" works without the author toggling
  per-cell lock state by hand?
- What exact `<sheetProtection>` attribute combination and cell-lock state does the target
  application require to honor sort/AutoFilter? Verify empirically against real files.
- How should `<protectedRanges>` (the OOXML "unlocked window inside a protected sheet" mechanism) be
  modeled in the public API?

Related: `sheet-protection-permits-requested-operations`, `sheet-protection-password-hash-compatibility`,
`workbook-structure-protection-authoring`, `cell-protection-locked-flag-and-sheet-protection`,
`autofilter-emits-filter-database-defined-name`.
