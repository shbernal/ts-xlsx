// The cell value model — the second bedrock primitive after addressing.
//
// A cell holds exactly one value, and that value's *type* is observable and drives
// everything downstream (serialization, number-format application, formula results).
// The honest shape here is a discriminated union: a value is either a JS primitive
// (null / number / string / boolean / Date) or one of the structural OOXML value
// shapes (error, formula, shared formula, rich text, hyperlink). There is no
// stringly-typed sentinel and no silent coercion between kinds — a numeric-looking
// string stays a string, because the caller's chosen type is the source of truth.

import type {Font} from './style.ts';

/** The observable kind of a cell's value. Both formula shapes report as `Formula`. */
export const ValueType = {
  Null: 'null',
  Number: 'number',
  String: 'string',
  Boolean: 'boolean',
  Date: 'date',
  Error: 'error',
  Formula: 'formula',
  RichText: 'richText',
  Hyperlink: 'hyperlink',
} as const;

export type ValueType = (typeof ValueType)[keyof typeof ValueType];

/** The canonical Excel error literals a cell (or formula result) can carry. */
export const ERROR_CODES = [
  '#N/A',
  '#REF!',
  '#NAME?',
  '#DIV/0!',
  '#NULL!',
  '#VALUE!',
  '#NUM!',
  '#SPILL!',
  '#CALC!',
  '#GETTING_DATA',
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

/** An in-cell error, e.g. `{error: '#REF!'}`. */
export interface ErrorValue {
  readonly error: ErrorCode;
}

/** One formatted run of a rich-text value. */
export interface RichTextRun {
  readonly text: string;
  readonly font?: Partial<Font>;
}

/** A value composed of independently-formatted text runs. */
export interface RichTextValue {
  readonly richText: readonly RichTextRun[];
}

/** A hyperlink cell: a URL plus the text (plain or rich) shown in the cell. */
export interface HyperlinkValue {
  readonly hyperlink: string;
  readonly text: string | RichTextValue;
  readonly tooltip?: string;
}

/** The cached result a formula carries — any scalar, a date, or an error. */
export type FormulaResult = number | string | boolean | Date | ErrorValue;

/** A cell whose value is computed by its own formula. */
export interface FormulaValue {
  readonly formula: string;
  readonly result?: FormulaResult;
}

/**
 * A cell that participates in a shared formula. The master cell carries the source
 * `sharedFormula` text plus the `ref` range it spans; slave cells reference the same
 * text with their own translated `result`.
 */
export interface SharedFormulaValue {
  readonly sharedFormula: string;
  readonly ref?: string;
  readonly result?: FormulaResult;
}

/** Everything a cell's value can be. `null` is the empty cell. */
export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | ErrorValue
  | FormulaValue
  | SharedFormulaValue
  | RichTextValue
  | HyperlinkValue;

const ERROR_SET: ReadonlySet<string> = new Set(ERROR_CODES);

export function isErrorValue(value: CellValue): value is ErrorValue {
  return typeof value === 'object' && value !== null && 'error' in value;
}

export function isFormulaValue(value: CellValue): value is FormulaValue {
  return typeof value === 'object' && value !== null && 'formula' in value;
}

export function isSharedFormulaValue(value: CellValue): value is SharedFormulaValue {
  return typeof value === 'object' && value !== null && 'sharedFormula' in value;
}

export function isRichTextValue(value: CellValue): value is RichTextValue {
  return typeof value === 'object' && value !== null && 'richText' in value;
}

export function isHyperlinkValue(value: CellValue): value is HyperlinkValue {
  return typeof value === 'object' && value !== null && 'hyperlink' in value;
}

/**
 * Classify a value into its observable {@link ValueType}. This is total over
 * {@link CellValue}: every legal value has exactly one type. A `Date` is a date even
 * when its time is `NaN` (an invalid date is still a date-typed cell); serialization,
 * not the model, decides what to do with it.
 */
export function detectValueType(value: CellValue): ValueType {
  if (value === null) return ValueType.Null;
  switch (typeof value) {
    case 'number':
      return ValueType.Number;
    case 'string':
      return ValueType.String;
    case 'boolean':
      return ValueType.Boolean;
    default:
      break;
  }
  if (value instanceof Date) return ValueType.Date;
  // Order matters: a hyperlink whose text is rich must classify as Hyperlink, and a
  // formula carrying a result must classify as Formula — check the outer shape first.
  if (isHyperlinkValue(value)) return ValueType.Hyperlink;
  if (isFormulaValue(value) || isSharedFormulaValue(value)) return ValueType.Formula;
  if (isRichTextValue(value)) return ValueType.RichText;
  if (isErrorValue(value)) return ValueType.Error;
  throw new TypeError(`unsupported cell value: ${describe(value)}`);
}

/** Whether a string is one of Excel's canonical error literals. */
export function isErrorCode(text: string): text is ErrorCode {
  return ERROR_SET.has(text);
}

function describe(value: unknown): string {
  if (typeof value === 'object' && value !== null) {
    return `object with keys [${Object.keys(value).join(', ')}]`;
  }
  return String(value);
}

/**
 * Normalise a raw assignment into a stored {@link CellValue}. `undefined` becomes the
 * empty cell (`null`); every other kind is validated by {@link detectValueType} and
 * passed through unchanged — the model never rewrites one value kind into another.
 *
 * @throws {TypeError} if the value is not a recognised cell-value shape.
 */
export function coerceCellValue(value: CellValue | undefined): CellValue {
  if (value === undefined) return null;
  // detectValueType throws on an unrecognised object shape, so this both validates
  // and gives a precise error at the assignment site rather than deep in serialization.
  detectValueType(value);
  return value;
}
