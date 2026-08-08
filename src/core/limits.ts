// Excel's limits on the grid's *geometry* — the other half of the bounds `address.ts` states.
//
// `MAX_ROW`/`MAX_COLUMN` bound where a cell can be; these bound how big a line can be drawn. The
// difference that matters is who enforces them: the addressing bounds are structural (a reference
// past XFD is not a reference), so the model refuses them outright, while these are *application*
// limits. The schema types `ht` and `width` as a bare `xsd:double` and puts no ceiling on either,
// so a file can carry more, and a reader that threw on one would be refusing to read a file Excel
// itself opens. Nothing here is enforced, therefore — these are for a caller sizing a line it is
// about to write, and the doc comments on `Row.height`/`Column.width` point here for that reason.
//
// Both numbers are what Excel Desktop accepts, measured rather than quoted: Microsoft's own
// specifications page rounds the row-height ceiling to "409 points", and Excel takes 409.5. See
// docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md for the probe.

/**
 * The tallest row Excel accepts, in points. Excel refuses 409.6 and takes 409.5, so a row asked to
 * hold more wrapped text than this cannot grow to fit it — beyond a few hundred points, laying such
 * a sheet out is also work Excel does lazily, leaving bands of the grid undrawn until clicked.
 */
export const MAX_ROW_HEIGHT = 409.5;

/**
 * The widest column Excel accepts, in character units of the workbook's default font. Excel refuses
 * 255.4, so unlike the row-height ceiling this one is exactly integral.
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
