// Decoding a worksheet cell's on-disk `<c>` payload into a model {@link CellValue}.
//
// This is the single value-decoding surface both readers share: the buffered reader
// (`./read.ts`) and the streaming row reader (`./read-rows.ts`). Keeping it in one place is
// what guarantees a cell read one row at a time decodes identically to the same cell read as
// part of a whole workbook — a divergence here would be a silent data bug in exactly one path.

import {isDateFormat, serialToDate} from '../../core/date.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {type CellValue, type FormulaResult, isErrorCode, type RichTextRun} from '../../core/value.ts';

/** The raw, still-textual pieces of a `<c>` element the SAX pass has gathered. */
export interface RawCell {
  /** The `t` attribute (`s`, `str`, `inlineStr`, `b`, `e`, `d`, or '' for a number). */
  readonly type: string;
  readonly hasFormula: boolean;
  readonly formula: string;
  readonly hasValue: boolean;
  readonly valueText: string;
  readonly inlineText: string;
  /** The formatted runs of a rich inline string, when the `<is>` held `<r>` elements rather than a
   * bare `<t>`. Absent (or empty) for a plain inline string, which decodes to `inlineText`. */
  readonly richTextRuns?: readonly RichTextRun[];
}

/**
 * Decode a gathered cell into its model value. A formula cell becomes a `{formula, result?}`
 * object (the on-disk `_xlfn.`/`_xlpm.` mangling stripped back to the readable name); a plain
 * numeric cell under a date number format becomes a {@link Date}; everything else decodes by its
 * `t` type. `numFmt` is the cell's resolved number-format code, used only for date detection.
 */
export function decodeCellContent(
  raw: RawCell,
  sharedStrings: readonly string[],
  numFmt: string | undefined
): CellValue {
  if (raw.hasFormula) {
    const stored = unmangleFunctions(raw.formula);
    const result = raw.hasValue ? decodeResult(raw.type, raw.valueText) : undefined;
    return result === undefined ? {formula: stored} : {formula: stored, result};
  }
  // An inline string built from `<r>` runs is rich text — surface its runs rather than flattening
  // them to the concatenated `inlineText` a plain string would decode to.
  if (raw.type === 'inlineStr' && raw.richTextRuns !== undefined && raw.richTextRuns.length > 0) {
    return {richText: raw.richTextRuns};
  }
  const value = decodeValue(raw.type, raw.valueText, raw.inlineText, raw.hasValue, sharedStrings);
  // A number stored under a date format is a date serial — surface it as a Date so a written
  // date round-trips as a date, not a bare number. Only plain numeric cells qualify; a string,
  // boolean, or formula result under a date format keeps its own kind.
  return typeof value === 'number' && numFmt !== undefined && isDateFormat(numFmt)
    ? serialToDate(value)
    : value;
}

function decodeValue(
  type: string,
  valueText: string,
  inlineText: string,
  hasValue: boolean,
  sharedStrings: readonly string[]
): CellValue {
  switch (type) {
    case 'inlineStr':
      return inlineText;
    case 'str':
      return valueText;
    case 'd':
      // A Strict-mode (ISO/IEC 29500 Strict) date cell stores an ISO 8601 value directly, not a
      // serial. Parse it literally — an ISO date is UTC — so it reads as the date it states rather
      // than a 1900-epoch serial the transitional decoder would fabricate from the text.
      return valueText === '' ? null : new Date(valueText);
    case 's': {
      const index = Number(valueText);
      return Number.isInteger(index) ? sharedStrings[index] ?? '' : '';
    }
    case 'b':
      return valueText === '1' || valueText === 'true';
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return hasValue ? Number(valueText) : null;
  }
}

/** Decode a formula's cached `<v>` result by its `t` type — shared by the buffered reader's shared-
 * formula clone resolution, which caches a result the same way a plain formula cell does. */
export function decodeFormulaResult(type: string, valueText: string): FormulaResult {
  return decodeResult(type, valueText);
}

function decodeResult(type: string, valueText: string): FormulaResult {
  switch (type) {
    case 'str':
      return valueText;
    case 'b':
      return valueText === '1' || valueText === 'true';
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return Number(valueText);
  }
}
