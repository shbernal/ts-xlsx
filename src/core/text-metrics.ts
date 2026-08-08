// How many lines a wrapped cell takes - the one measurement a writer needs and the format does
// not record.
//
// A row that states no height is auto-fitted by Excel at paint time, and auto-fitting a sheet
// whose columns all wrap and whose cells run to a couple of thousand characters is work Excel does
// lazily and incompletely: bands open blank until they are clicked. Writing an explicit height
// settles the geometry before the file is opened, and to write one you have to know the line count.
//
// This counts *characters against a character-unit width*. It does not measure glyphs, and it is
// not the deferred font-metric question that
// docs/knowledge/specs/default-font-must-not-be-assumed-for-column-widths.md leaves open: a column
// width is already expressed in character units, so counting characters is dimensionally honest
// and needs no metric table. It is also, for the same reason, an approximation for any
// proportional face - `WWW` and `iii` are one character unit each here and are not on screen.

/**
 * The number of lines `text` occupies when wrapped at `width` character units - the width unit a
 * column states, so `estimateWrappedLines(cell, sheet.getColumn(2).width ?? 8.43)` is the shape of
 * the call.
 *
 * A hard break opens a line of its own and what follows wraps independently, matching how Excel
 * lays a wrapped cell out. The empty string is one line, not zero: a cell always occupies its row.
 *
 * An estimate, and only ever that. It counts characters, so it is exact for a monospaced face and
 * approximate for every other - a run of `W`s wraps sooner on screen than this predicts, a run of
 * `i`s later. Excel's own auto-fit measures glyphs; this exists so that a writer can state *a*
 * height rather than leave the sheet to be laid out lazily on open, and being within a line of the
 * truth is what that needs.
 *
 * @throws {RangeError} if `width` is not a positive finite number - a column of zero width wraps
 *   nothing, and silently answering `Infinity` or `NaN` would put that straight into a row height.
 */
export function estimateWrappedLines(text: string, width: number): number {
  if (!Number.isFinite(width) || width <= 0) {
    throw new RangeError(`wrap width ${width} is not a positive finite number of character units`);
  }
  let lines = 0;
  // A lone \r is not a break Excel writes, but a value assembled by a caller on Windows may carry
  // one; counting it as a break costs nothing and mis-counting it as a character would show up as
  // a row one line short.
  for (const segment of text.split(/\r\n?|\n/)) {
    lines += Math.max(1, Math.ceil(segment.length / width));
  }
  return lines;
}
