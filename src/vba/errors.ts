/**
 * Thrown when a VBA project (`vbaProject.bin`) is present but cannot be parsed — a malformed
 * compound file, a corrupt compressed stream, or a `dir` record that does not conform to [MS-OVBA].
 * A workbook with no macros never produces this: {@link Workbook.vbaProject} is `undefined` instead.
 *
 * The parser treats the blob as hostile input (a spreadsheet library parses untrusted files), so every
 * malformed structure fails closed with this error rather than crashing, hanging, or over-allocating.
 */
export class VbaParseError extends Error {
  override readonly name = 'VbaParseError';
}
