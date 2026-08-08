# Grid limits

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `MAX_COLUMN_WIDTH`

<sub>const</sub>

The widest column Excel accepts, in character units of the workbook's default font. Excel refuses
255.4, so unlike the row-height ceiling this one is exactly integral.

Character units, not points or pixels: a width is a count of digits of the default font's
*maximum digit width*, which is why there is no companion `DEFAULT_COLUMN_WIDTH` constant here.
The width a column takes when it states none is a function of that font — the familiar 8.43 holds
for Calibri 11 and not for a workbook whose normal style says otherwise (Excel reports 8.09 for
Aptos Narrow 11). `sheet.properties.defaultColWidth` is what a file declares, and
docs/knowledge/specs/default-font-must-not-be-assumed-for-column-widths.md is why assuming a
value for it is a bug rather than a shortcut.

```ts
const MAX_COLUMN_WIDTH: 255
```

---

### `MAX_ROW_HEIGHT`

<sub>const</sub>

The tallest row Excel accepts, in points. Excel refuses 409.6 and takes 409.5, so a row asked to
hold more wrapped text than this cannot grow to fit it — beyond a few hundred points, laying such
a sheet out is also work Excel does lazily, leaving bands of the grid undrawn until clicked.

```ts
const MAX_ROW_HEIGHT: 409.5
```
