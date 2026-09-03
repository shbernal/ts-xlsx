// What a CSV delimiter may be, for both halves of the codec.
//
// The reader refused a multi-character delimiter and the writer accepted anything, in a module whose
// header is entirely about lossless round-tripping. So `readCsv(writeCsv(wb, {delimiter: "||"}))`
// threw on text this same codec had just produced, and `writeCsv({delimiter: ""})` was worse than a
// throw: `field.includes("")` is true of every field, so every field was quoted and the output had no
// separators in it at all -- a file that parses as one column and loses nothing visibly.
//
// One validator, called from both entry points, is the only arrangement in which the two halves
// cannot disagree about what a delimiter is.

import {quoted} from '../../errors.ts';

/**
 * Refuse a delimiter the codec cannot honour.
 *
 * A `RangeError` rather than an `AuthoringError`: this is a single scalar out of range, which
 * `src/errors.ts` reserves for the native types on the grounds that they exist for exactly that.
 *
 * @throws {RangeError} if the delimiter is not exactly one character.
 */
export function assertDelimiter(delimiter: string): void {
  if (delimiter.length !== 1) {
    throw new RangeError(`CSV delimiter must be a single character, got ${quoted(delimiter)}`);
  }
}
