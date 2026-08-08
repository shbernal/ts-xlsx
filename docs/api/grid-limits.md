# Grid limits

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `MAX_COLUMN_WIDTH`

<sub>const</sub>

The widest column Excel accepts being set to, in character units of the workbook's default font.
Excel refuses 255.4, so unlike the row-height ceiling this one is exactly integral.

It is also the weaker of the two ceilings: it binds assignment only, and not a file at all.
Excel honours a `width` of 1000 read from a package, renders the column at it, and round-trips
it verbatim through its own save — where an over-limit row height is clamped away. So a width
above this is a column no Excel user could have produced by dragging, not a value at risk.

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

The tallest row Excel accepts being set to, in points: it refuses 409.6 and takes 409.5. A row
asked to hold more wrapped text than this cannot grow to fit it, and the overflow is simply not
shown.

A file may state more, and stating more loses the value rather than the file. Excel opens such a
package without complaint and silently clamps the row — to 409.6, a tick *above* what it lets
you assign, being 8192 twentieths of a point and so the width of the field it is read into — and
writes 409.6 back on its next save. Check against this constant to keep a stated height from
quietly becoming a different one.

```ts
const MAX_ROW_HEIGHT: 409.5
```
