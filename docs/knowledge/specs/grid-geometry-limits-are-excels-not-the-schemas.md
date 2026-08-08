# A row's height and a column's width are bounded by Excel, not by the schema

Cluster: worksheet

## Scenario

A caller sizing a line it is about to write - a header band, a row wrapping a long paragraph -
needs to know how large it may go. The format is no help: `CT_Row@ht` and `CT_Col@width` are both
a bare `xsd:double` in `sml.xsd`, with no `maxInclusive` on either. The ceiling is Excel's, and a
row asked to hold more wrapped text than Excel can draw does not clip gracefully - the layout has
no solution, and bands of the sheet open blank until they are clicked.

Microsoft's published *Excel specifications and limits* table states "Row height: 409 points" and
"Column width: 255 characters". The first of those is rounded.

## Measured behavior

Excel Desktop (Microsoft 365, Windows, workbook default font Aptos Narrow 11), driven over COM
against a blank workbook - assign, then read back:

| Assigned `RowHeight` | Result |
| --- | --- |
| 409 | accepted, reads back 409 |
| 409.4 | accepted, reads back 409.4 |
| **409.5** | **accepted, reads back 409.5** |
| 409.6 | refused - "Unable to set the RowHeight property of the Range class" |
| 410, 500 | refused |

| Assigned `ColumnWidth` | Result |
| --- | --- |
| **255** | **accepted, reads back 255** |
| 255.4, 255.5, 255.9 | refused |
| 256, 300 | refused |

So the row-height ceiling is 409.5 and the published 409 is a rounding of it; the column-width
ceiling is exactly integral at 255, which is why one constant carries a fraction and the other
does not. `MAX_ROW_HEIGHT` and `MAX_COLUMN_WIDTH` (`src/core/limits.ts`) hold these, and
`src/core/limits.test.ts` pins them with a comment pointing back here - a future reader who
"corrects" 409.5 to 409 against the specifications page fails that test and finds this note.

The same probe recorded `StandardWidth = 8.09` and `StandardHeight = 14.5` for that workbook,
against the familiar 8.43 / 15 of a Calibri-11 one. That is the second fact here, and the reason
no `DEFAULT_COLUMN_WIDTH` constant sits beside the two ceilings: the default width is a function
of the workbook's default font, not a property of the format. See
`default-font-must-not-be-assumed-for-column-widths`.

## Why nothing enforces them

`Row.height` and `Column.width` accept values above these ceilings, deliberately. The setters are
not an authoring-only surface: `read-worksheet.ts` loads a foreign `ht`/`width` straight through
them, so a bound that threw would refuse to read files Excel itself opens without complaint. The
constants are published so a caller can check before writing, and the two setters' doc comments
point at them.

That leaves an authoring mistake (a 5000-point row) undetected by the library. The alternative -
an internal restore path for the reader so the public setter can validate - buys a diagnostic at
the cost of a second way to set the same property, and is worth revisiting only if the mistake
turns out to be common.

## Provenance

`source: excel-desktop-verification`. One Excel build, one host, state-observable via COM
assignment - the tier that ADR-0013 describes. Not run in CI and not reproducible without Excel;
the numbers are locked by the unit test instead.
