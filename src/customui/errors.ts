/**
 * Thrown when a `customUI` ribbon-customisation part (`customUI/customUI.xml` or
 * `customUI/customUI14.xml`) is present but cannot be parsed into a {@link CustomUiDocument} — malformed
 * XML, a missing or unrecognised `<customUI>` root namespace, or nesting deep enough to look hostile.
 * A workbook with no ribbon customisation never produces this: {@link Workbook.customUI} is an empty
 * array instead.
 *
 * The parser treats the part as hostile input (a spreadsheet library parses untrusted files), so a
 * malformed structure fails closed with this error rather than yielding a half-built tree.
 */
export class CustomUiParseError extends Error {
  override readonly name = 'CustomUiParseError';
}
