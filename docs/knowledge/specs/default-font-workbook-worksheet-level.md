# A document-wide default font must be settable once, per workbook (and per worksheet)

Cluster: styles

## Scenario

A user building a workbook wants every cell to inherit a chosen base font (family, size, color)
instead of the hard-coded Calibri 11pt default, without setting the font on every row, column, or
cell. They expect to declare a default font **once** — per workbook, with an optional per-worksheet
override — and have empty and populated cells alike render in it. A correctness symptom motivates
this too: when the written package carries an incomplete or missing default font entry in the styles
part, some applications (Apple Numbers, and Excel in some cases) warn "missing default font" on
open, because empty cells fall back to a default the styles part does not properly define.

## Desired behavior

- Expose a first-class way to set the document's default font once, rather than forcing
  per-cell/row/column styling. Two natural levels:
  - **Workbook-level default font** — the base font applied to all sheets unless overridden. This is
    the first entry of the styles part's font table, referenced by the "Normal" cell style; changing
    it changes the rendered font of every cell with no explicit font override, **including empty
    cells**.
  - **Optional worksheet-level override** — e.g. a default/style option when adding a worksheet
    (`addWorksheet(name, { style: { font: { size, name } } })`-shaped).
- **Correctness invariant to lock once implemented:** the written styles part always contains a
  well-formed default font entry (name, size, family, color/theme, scheme) referenced by the base
  cell style, so foreign readers do not warn about a missing default font for empty cells. This
  holds for the built-in Calibri 11 default and especially for a custom default.
- **Round-trip:** reading a file with a non-Calibri default font surfaces that default through the
  same API used to set it, so read-modify-write preserves it.

## Prior art / workarounds

- Post-processing every cell after writing to stamp `font.size`/`font.name` on any missing them —
  O(cells) and misses truly-empty cells unless iterating with `includeEmpty`.
- Monkey-patching the styles serializer to inject a font entry (fragile, reaches into internals).
- Wrapping `addRow` to stamp a default font on each new row.

All are brute-force symptoms of there being no supported knob for the base font.

## Open questions

- Precedence/merge: workbook default vs worksheet default vs column/row default vs explicit cell
  font — confirm the inheritance chain, and whether partial fonts merge (set only size, keep default
  name) or replace wholesale.
- Does setting a workbook default rewrite the theme's minor/major font, or only the styles font
  table? (Affects whether Excel labels it a theme font vs a direct font.)
- Interaction with the shared-strings/inline default and with rich-text runs carrying their own
  fonts.

Related: `default-font-must-not-be-assumed-for-column-widths` (the geometry face of the same
default-font question), `theme-color-font-backed-by-theme-part`, `template-styles-survive-read-write-roundtrip`.
