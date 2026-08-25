# Sheet-view boolean flags, and showFormulas is not formula-bar visibility

Cluster: styles

## Scenario

A user configuring a worksheet's on-screen view wants to know which visibility toggles the view
object exposes. They see booleans for gridlines, row and column headers, and rulers, and ask whether
a "show formula bar" toggle also exists. The confusion is between two distinct concepts: the
sheet-view attribute that makes cells display their **formula text** instead of their computed
result, and the application-level chrome control for whether the **formula bar** is visible. The
former is a legitimate per-sheet-view attribute; the latter is not a sheet-view attribute at all and
cannot be expressed per worksheet.

## Desired behavior

The worksheet view surface should expose the set of sheet-view boolean attributes the OOXML
`sheetView` element actually supports, with names that faithfully reflect their meaning, and
should **not** invent a per-sheet toggle for formula-bar visibility that has no representation in
the file format.

## Format facts

- The OOXML `sheetView` element carries per-view booleans including `showGridLines`,
  `showRowColHeaders`, `showRuler`, `showFormulas` (display formula text in cells rather than
  results), `showZeros`, `showOutlineSymbols`, `rightToLeft`, `tabSelected`, and
  `defaultGridColor`. It also carries view mode (normal, pageBreakPreview, pageLayout),
  `zoomScale`, selection, and pane and freeze state.
- There is **no** `showFormulaBar` attribute on `sheetView`. Formula-bar visibility is an
  application or window-level preference, not a per-worksheet-view setting, so it cannot be
  round-tripped per sheet.
- `showFormulas` is easy to mistake for a "formula bar" toggle because of the similar name, but
  its effect is to render each cell's formula source in the grid instead of the evaluated value.

## Open questions

- Which `sheetView` booleans should be first-class typed properties on the view object, and which
  passed through generically? At minimum `showGridLines`, `showRowColHeaders`, `showRuler`,
  `showFormulas`, `showZeros`, and `rightToLeft` are common enough to type explicitly.
- Should the API document the `showFormulas` semantic clearly, formula text in cells rather than
  formula-bar chrome, to head off this exact confusion, via a doc comment or a more explicit name?
- If a caller genuinely wants formula-bar chrome hidden, that belongs to a workbook-level
  window or application-settings surface if we model it at all, and it should not be attached to a
  worksheet view.
