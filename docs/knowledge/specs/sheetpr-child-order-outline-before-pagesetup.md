# The sheetPr child order: outlinePr must precede pageSetUpPr

Cluster: xlsx-io

## Scenario

A worksheet is configured with both page-setup fit-to-page scaling and outline summary properties
(e.g. summary rows placed above their detail rather than below). The worksheet's sheet-properties
element (`<sheetPr>`) must emit its outline-properties child (`<outlinePr>`) before its
page-setup-properties child (`<pageSetUpPr>`), the order the CT_SheetPr schema sequence requires. If
they are emitted in the wrong order the file is rejected as corrupt by Excel, which offers only to
recover the workbook and then drops the affected sheet. Setting either property alone produces a
valid file; the corruption arises only from their combination — a sign of a child-ordering bug.

> Spec note, not a corpus case: probing the current writer shows it does **not** emit an `<outlinePr>`
> child from the outline-summary settings at all (only `<pageSetUpPr fitToPage="1"/>` appears), so
> the bad ordering is not currently reachable through the authoring API and there is no failing
> serialization to assert. The durable value is the schema-order requirement, to guard once outline
> summary authoring emits `<outlinePr>`.

## Desired behavior

- Writing a worksheet with **both** fit-to-page scaling and an outline summary property produces a
  valid, non-corrupt worksheet part.
- Within `<sheetPr>`, the children follow the CT_SheetPr sequence — in particular **`<outlinePr>`
  before `<pageSetUpPr>`** (the full order is `tabColor`, `outlinePr`, `pageSetUpPr`).
- The fit-to-page setting round-trips (the `<pageSetUpPr>` records `fitToPage`) alongside the outline
  summary setting.
- **The summary-placement settings are precisely typed on the public worksheet-properties surface.**
  The worksheet properties type must fully type an `outlineProperties` field with its two booleans —
  one for whether outline summary rows appear *below* their detail rows (vs. above), one for whether
  summary columns appear to the *right* of their detail columns (vs. left) — which map to the
  `<outlinePr>` attributes `summaryBelow`/`summaryRight`. Reports of the runtime honoring these values
  while the public TypeScript declaration omitted `outlineProperties` (forcing an untyped index-access
  workaround) are the type-surface face of the same outline feature: the types *are* the docs, so a
  supported outline property that is settable at runtime but absent from the type is a defect. Absence
  of the setting corresponds to the format default (summary below and to the right).

## Open questions

- Authoring gap: the outline summary properties (`summaryBelow`/`summaryRight`) currently do not
  surface as an `<outlinePr>` child — is that a missing write path, or are they carried elsewhere?
  This must be settled before the ordering guard can be exercised.
- Once `<outlinePr>` is emitted, a corpus case can assert the child order via the worksheet XML (the
  same element-order technique the `comment-and-table-coexist-on-same-sheet` case uses for the
  drawing/legacyDrawing/tableParts sequence).

Related: `comment-and-table-coexist-on-same-sheet`, `row-outline-collapsed-flag-belongs-on-summary-row`,
`column-width-and-pagesetup-roundtrip-fidelity`.
