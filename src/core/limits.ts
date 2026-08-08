// Excel's limits on the grid's *geometry* — the other half of the bounds `address.ts` states.
//
// `MAX_ROW`/`MAX_COLUMN` bound where a cell can be; these bound how big a line can be *set*. The
// difference that matters is who enforces them: the addressing bounds are structural (a reference
// past XFD is not a reference), so the model refuses them outright, while these are limits on
// *assignment* in Excel's own UI and object model. They are not limits on what a package may
// carry. The schema types `ht` and `width` as a bare `xsd:double` with no ceiling on either, and
// Excel opens an over-limit file clean — no repair prompt, no repair log — so a reader that threw
// on one would refuse a file Excel accepts. Nothing here is enforced, therefore, on either the
// read or the write path; these are for a caller that wants the size it states to be the size
// Excel uses, and the doc comments on `Row.height`/`Column.width` point here for that reason.
//
// What Excel does with an over-limit file is not symmetric, and the difference is why only one of
// these two numbers describes a value that survives: a row is silently clamped, a column is not.
// Both ceilings and both file behaviours were measured rather than quoted — Microsoft's own
// specifications page rounds the row-height ceiling to "409 points", and Excel takes 409.5. See
// docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md for the probes.

/**
 * The tallest row Excel accepts being set to, in points: it refuses 409.6 and takes 409.5. A row
 * asked to hold more wrapped text than this cannot grow to fit it, and the overflow is simply not
 * shown.
 *
 * A file may state more, and stating more loses the value rather than the file. Excel opens such a
 * package without complaint and silently clamps the row — to 409.6, a tick *above* what it lets
 * you assign, being 8192 twentieths of a point and so the width of the field it is read into — and
 * writes 409.6 back on its next save. Check against this constant to keep a stated height from
 * quietly becoming a different one.
 */
export const MAX_ROW_HEIGHT = 409.5;

/**
 * The widest column Excel accepts being set to, in character units of the workbook's default font.
 * Excel refuses 255.4, so unlike the row-height ceiling this one is exactly integral.
 *
 * It is also the weaker of the two ceilings: it binds assignment only, and not a file at all.
 * Excel honours a `width` of 1000 read from a package, renders the column at it, and round-trips
 * it verbatim through its own save — where an over-limit row height is clamped away. So a width
 * above this is a column no Excel user could have produced by dragging, not a value at risk.
 *
 * Character units, not points or pixels: a width is a count of digits of the default font's
 * *maximum digit width*, which is why there is no companion `DEFAULT_COLUMN_WIDTH` constant here.
 * The width a column takes when it states none is a function of that font — the familiar 8.43 holds
 * for Calibri 11 and not for a workbook whose normal style says otherwise (Excel reports 8.09 for
 * Aptos Narrow 11). `sheet.properties.defaultColWidth` is what a file declares, and
 * docs/knowledge/specs/default-font-must-not-be-assumed-for-column-widths.md is why assuming a
 * value for it is a bug rather than a shortcut.
 */
export const MAX_COLUMN_WIDTH = 255;
