// How many lines a wrapped cell takes - the one measurement a writer needs and the format does
// not record.
//
// A row that states no height records no geometry, so the height is whoever opens the file's
// answer. Excel Desktop's answer is an auto-fit computed on open - promptly and correctly, but
// saturating at MAX_ROW_HEIGHT, so past ~28 lines it stops answering the question - and no answer
// at all is what a consumer that does not implement that auto-fit gets, this library's reader
// included. Stating a height is how the geometry stops depending on the reader, and to state one
// you have to know the line count. Measured in
// docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md, which also buries the
// claim this comment used to make, that such a sheet opens with blank unpainted bands. It does not.
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
 * An estimate, and only ever that. It counts characters, so it is exact for a monospaced face that
 * wraps mid-word and approximate for every other - a run of `W`s wraps sooner on screen than this
 * predicts, a run of `i`s later. Against Excel it reads a shade *low*, because Excel breaks at word
 * boundaries and its usable width is about 0.64 character units under the stated one: measured at
 * 5 lines where Excel laid out 6, 25 where Excel laid out 26. This exists so that a writer can
 * state *a* height rather than leave one to the application that opens the file, and being within a
 * line of the truth is what that needs.
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
