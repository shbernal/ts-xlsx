import {XlsxError} from '../../errors.ts';

/**
 * Thrown when an `.xlsb` package's binary BIFF12 content cannot be parsed — a record whose declared
 * size runs past the end of its part, a truncated record header, or a structure that does not conform
 * to [MS-XLSB].
 *
 * The record streams are hostile input (a spreadsheet library parses untrusted files), so a malformed
 * part fails closed with this error rather than crashing, hanging, or over-allocating. It is distinct
 * from {@link UnsupportedFormatError}, which reports that the *container* is not a format we read at
 * all; by the time this is raised the input has already been recognised as an `.xlsb`.
 */
export class XlsbParseError extends XlsxError {
  override readonly name = 'XlsbParseError';
  override readonly code = 'malformed-input';
}
