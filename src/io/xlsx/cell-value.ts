// Decoding a worksheet cell's on-disk `<c>` payload into a model {@link CellValue}.
//
// This is the single value-decoding surface both readers share: the buffered reader
// (`./read.ts`) and the streaming row reader (`./read-rows.ts`). Keeping it in one place is
// what guarantees a cell read one row at a time decodes identically to the same cell read as
// part of a whole workbook — a divergence here would be a silent data bug in exactly one path.

import {isDateFormat, serialToDate} from '../../core/date.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {
  type CellValue,
  type FormulaResult,
  isErrorCode,
  type RichTextRun,
  type RichTextValue,
} from '../../core/value.ts';
import {boolStrict} from './xml-read.ts';

/**
 * One entry of the shared-strings pool. A `<si>` built from a bare `<t>` is a plain string; a `<si>`
 * built from `<r>` runs is rich text — so a `t="s"` cell can resolve to either kind, and rich text
 * that Excel pooled reads back with its per-run formatting intact rather than flattened to text.
 */
export type SharedString = string | RichTextValue;

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
  sharedStrings: readonly SharedString[],
  numFmt: string | undefined,
): CellValue {
  if (raw.hasFormula) {
    const stored = unmangleFunctions(raw.formula);
    const result = raw.hasValue ? decodeFormulaResult(raw.type, raw.valueText, numFmt) : undefined;
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
  sharedStrings: readonly SharedString[],
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
      // A `t="s"` cell indexes the shared pool; the entry is a plain string or, when Excel pooled a
      // rich value, a {@link RichTextValue} whose runs surface here rather than being flattened.
      const index = Number(valueText);
      return Number.isInteger(index) ? (sharedStrings[index] ?? '') : '';
    }
    case 'b':
      return boolStrict(valueText);
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return hasValue ? Number(valueText) : null;
  }
}

/** Decode a formula's cached `<v>` result by its `t` type, coercing a numeric result under a date
 * `numFmt` to a {@link Date} exactly as a bare numeric cell is — so a date-valued formula result
 * (e.g. `TODAY()`) reads back as a Date, not a serial. Shared by the buffered reader's shared-formula
 * clone resolution, which caches a result the same way a plain formula cell does. */
export function decodeFormulaResult(
  type: string,
  valueText: string,
  numFmt?: string,
): FormulaResult {
  const result = decodeResult(type, valueText);
  return typeof result === 'number' && numFmt !== undefined && isDateFormat(numFmt)
    ? serialToDate(result)
    : result;
}

// The formula-result subset of `decodeValue`: a cached result is only ever a string, boolean,
// error, or number — never a shared-string index, inline string, or Strict-mode date — so this
// handles just those cases rather than the full cell-value grammar.
function decodeResult(type: string, valueText: string): FormulaResult {
  switch (type) {
    case 'str':
      return valueText;
    case 'b':
      return boolStrict(valueText);
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return Number(valueText);
  }
}
