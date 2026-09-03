// Decoding a worksheet cell's on-disk `<c>` payload into a model {@link CellValue}.
//
// This is the single value-decoding surface both readers share: the buffered reader
// (`./read.ts`) and the streaming row reader (`./read-rows.ts`). Keeping it in one place is
// what guarantees a cell read one row at a time decodes identically to the same cell read as
// part of a whole workbook. A divergence here would be a silent data bug in exactly one path.

import {coerceDateSerial, type DateEpoch, parseDateText} from '../../core/date.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {
  type CellValue,
  type FormulaResult,
  isErrorCode,
  type RichTextRun,
  type RichTextValue,
} from '../../core/value.ts';
import {decodeSpreadsheetText, numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {boolStrict} from '../../xml/xml-scan.ts';

/**
 * One entry of the shared-strings pool. A `<si>` built from a bare `<t>` is a plain string; a `<si>`
 * built from `<r>` runs is rich text, so a `t="s"` cell can resolve to either kind, and rich text
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
 * `t` type. `numFmt` is the cell's resolved number-format code, used only for date detection, and
 * `epoch` the workbook's date system, which is what a serial under such a format counts from.
 */
export function decodeCellContent(
  raw: RawCell,
  sharedStrings: readonly SharedString[],
  numFmt: string | undefined,
  epoch: DateEpoch,
): CellValue {
  if (raw.hasFormula) {
    const stored = unmangleFunctions(raw.formula);
    const result = raw.hasValue
      ? decodeFormulaResult(raw.type, raw.valueText, numFmt, epoch)
      : undefined;
    return result === undefined ? {formula: stored} : {formula: stored, result};
  }
  // An inline string built from `<r>` runs is rich text: surface its runs rather than flattening
  // them to the concatenated `inlineText` a plain string would decode to.
  if (raw.type === 'inlineStr' && raw.richTextRuns !== undefined && raw.richTextRuns.length > 0) {
    return {richText: raw.richTextRuns};
  }
  const value = decodeValue(raw.type, raw.valueText, raw.inlineText, raw.hasValue, sharedStrings);
  // A number stored under a date format is a date serial: surface it as a Date so a written date
  // round-trips as a date, not a bare number.
  return coerceDateSerial(value, numFmt, epoch);
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
      // Already decoded per `<t>` as it was gathered; the accumulator owns that seam.
      return inlineText;
    case 'str':
      return decodeSpreadsheetText(valueText);
    case 'd':
      // A Strict-mode (ISO/IEC 29500 Strict) date cell stores an ISO 8601 value directly, not a
      // serial. Parse it literally, since an ISO date is UTC, so it reads as the date it states rather
      // than a 1900-epoch serial the transitional decoder would fabricate from the text.
      return parseDateText(valueText);
    case 's': {
      // A `t="s"` cell indexes the shared pool; the entry is a plain string or, when Excel pooled a
      // rich value, a {@link RichTextValue} whose runs surface here rather than being flattened.
      //
      // Read through the integer grammar rather than a bare `Number()`, for the reason the numeric
      // default branch below spells out: `Number('')` is 0 and 0 is an integer, so a present-but-empty
      // `<v/>` would resolve to the *first* pooled string: a wrong value, not a missing one.
      const index = numInteger(valueText, 0);
      return index === undefined ? '' : (sharedStrings[index] ?? '');
    }
    case 'b':
      return boolStrict(valueText);
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default: {
      if (!hasValue) return null;
      // The same grammar the numeric *attributes* read through, for the same reason: a bare
      // `Number()` turns a token it cannot parse into `NaN`, which is a number and so satisfies
      // every guard downstream. The writer already refuses to emit a `<v>` for a non-finite number,
      // so `NaN` was only ever a claim the model held for one step before the next write dropped it;
      // reading it as no value reaches the same file with no lie in the middle. A blank `<v>` goes
      // the same way rather than becoming `Number('')`'s zero: Excel writes no `<v>` at all for an
      // empty cell, so a present-but-empty one is malformed, and "nothing" is the honest reading.
      return numFinite(valueText) ?? null;
    }
  }
}

/** Decode a formula's cached `<v>` result by its `t` type, coercing a numeric result under a date
 * `numFmt` to a {@link Date} exactly as a bare numeric cell is, so a date-valued formula result
 * (e.g. `TODAY()`) reads back as a Date, not a serial. Shared by the buffered reader's shared-formula
 * clone resolution, which caches a result the same way a plain formula cell does. */
export function decodeFormulaResult(
  type: string,
  valueText: string,
  numFmt: string | undefined,
  epoch: DateEpoch,
): FormulaResult | undefined {
  return coerceDateSerial(decodeResult(type, valueText), numFmt, epoch);
}

// The formula-result subset of `decodeValue`: a cached result is only ever a string, boolean,
// error, or number, never a shared-string index, inline string, or Strict-mode date, so this
// handles just those cases rather than the full cell-value grammar.
function decodeResult(type: string, valueText: string): FormulaResult | undefined {
  switch (type) {
    case 'str':
      // The cached result of a string formula is a cell value, and Excel escapes and decodes it as
      // one. Verified on this host, a `<v>` of `_x0041_` under `t="str"` reads back as `A`. The
      // sibling `<v>` types are not text: a number, a boolean, and an error code have no `_` in
      // their grammars, so only this branch decodes.
      return decodeSpreadsheetText(valueText);
    case 'b':
      return boolStrict(valueText);
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      // Narrowed exactly as a plain numeric cell is: an unparseable cached result is no cached
      // result, which is the state the cell would have reached anyway on the next write.
      return numFinite(valueText);
  }
}
