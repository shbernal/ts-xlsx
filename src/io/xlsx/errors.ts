import {XlsxError} from '../../errors.ts';

/**
 * Thrown when an `.xlsx` package's XML content is well-formed but does not describe a workbook this
 * library can act on — a `xl/workbook.xml` that declares no worksheets, say.
 *
 * It sits between two neighbours. {@link XmlParseError} reports that the *markup* did not parse;
 * {@link UnsupportedFormatError} reports that the *container* is not one we read at all. This one
 * fires when both of those were fine and the document itself is nonetheless not a workbook.
 *
 * The reader is deliberately lenient about content it does not recognise — an unknown element is
 * skipped, not fatal — so this is rare by design, and reaching it means something a workbook cannot
 * do without being corrupt.
 */
export class XlsxParseError extends XlsxError {
  override readonly name = 'XlsxParseError';
  override readonly code = 'malformed-input';
}
