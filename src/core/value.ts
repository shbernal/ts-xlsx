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
  /** The clickable extent (`'D1:H1'`) when the link spans a range whose top-left corner is this
   * cell. Absent for an ordinary single-cell link. The destination and label live on the top-left
   * cell; `range` records how far Excel highlights the clickable area so it survives a round-trip. */
  readonly range?: string;
}

/** The cached result a formula carries — any scalar, a date, or an error. */
export type FormulaResult = number | string | boolean | Date | ErrorValue;

/** A cell whose value is computed by its own formula. */
export interface FormulaValue {
  readonly formula: string;
  readonly result?: FormulaResult;
}

/**
 * A cell that participates in a shared formula — a clone of a master formula cell filled across a
 * range. `sharedFormula` is the master cell's address (e.g. `'B1'`); the master itself is a plain
 * {@link FormulaValue}. On read, the clone's own formula is the master's translated to the clone's
 * position and `result` is the clone's cached value; on write, the clones of a master collapse into
 * OOXML's shared-formula grouping.
 */
export interface SharedFormulaValue {
  readonly sharedFormula: string;
  /** The master's formula translated to this cell's position. Filled in on read; a clone assigned by
   * a caller carries only `sharedFormula`, and the writer recovers the formula from the master. */
  readonly formula?: string;
  readonly result?: FormulaResult;
}

/**
 * A cell computed by a What-If-Analysis data table (`<f t="dataTable">`) — the OOXML formula kind that
 * fills a range by re-evaluating a model against a grid of substituted input cells. The library does
 * not evaluate it; it preserves the declaration so a read-modify-write cycle re-emits it verbatim
 * rather than silently dropping the data-table kind.
 */
export interface DataTableFormulaValue {
  readonly shareType: 'dataTable';
  /** The range the data table fills, e.g. `'B2:B5'`. */
  readonly ref: string;
  /** Whether the table substitutes two inputs (a 2-D data table) rather than one. */
  readonly dataTable2D?: boolean;
  /** For a 1-D table, whether the input runs along the row rather than down the column. */
  readonly dataTableRow?: boolean;
  /** The first (row) input-cell reference. */
  readonly r1?: string;
  /** The second (column) input-cell reference, present for a 2-D table. */
  readonly r2?: string;
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
  | DataTableFormulaValue
  | RichTextValue
  | HyperlinkValue;

const ERROR_SET: ReadonlySet<string> = new Set(ERROR_CODES);

// Whether a value is a non-null object carrying `key` — the object-shaped {@link CellValue} kinds are all
// discriminated by the presence of a single property, so every guard below narrows through this one test.
function hasKey<K extends string>(value: unknown, key: K): value is Record<K, unknown> & object {
  return typeof value === 'object' && value !== null && key in value;
}

export function isErrorValue(value: CellValue): value is ErrorValue {
  return hasKey(value, 'error');
}

export function isFormulaValue(value: CellValue): value is FormulaValue {
  // A shared-formula clone resolved on read carries both its master address (`sharedFormula`) and the
  // translated `formula`; it is a SharedFormulaValue, so exclude it here to keep the two kinds distinct.
  return hasKey(value, 'formula') && !('sharedFormula' in value);
}

export function isSharedFormulaValue(value: CellValue): value is SharedFormulaValue {
  return hasKey(value, 'sharedFormula');
}

export function isDataTableFormulaValue(value: CellValue): value is DataTableFormulaValue {
  return hasKey(value, 'shareType') && value.shareType === 'dataTable';
}

export function isRichTextValue(value: CellValue): value is RichTextValue {
  return hasKey(value, 'richText');
}

export function isHyperlinkValue(value: CellValue): value is HyperlinkValue {
  return hasKey(value, 'hyperlink');
}

/**
 * Flatten a rich-text value to its plain text by concatenating every run's text in order. This is the
 * text a consumer that cannot render per-run formatting (a CSV field, a pivot cache entry) sees, and
 * the string a rich cell reads as when its formatting is discarded.
 */
export function richTextToPlain(value: RichTextValue): string {
  return value.richText.map((run) => run.text).join('');
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
  if (isFormulaValue(value) || isSharedFormulaValue(value) || isDataTableFormulaValue(value)) {
    return ValueType.Formula;
  }
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
 * empty cell (`null`); every other kind is validated by {@link detectValueType}. The
 * model never rewrites one value *kind* into another (a numeric-looking string stays a
 * string) — the single exception is formula text, which is canonicalised to the OOXML
 * stored form (no leading `=`) so round-trips are idempotent regardless of how the
 * caller supplied it.
 *
 * @throws {TypeError} if the value is not a recognised cell-value shape.
 */
export function coerceCellValue(value: CellValue | undefined): CellValue {
  if (value === undefined) return null;
  // detectValueType throws on an unrecognised object shape, so this both validates
  // and gives a precise error at the assignment site rather than deep in serialization.
  detectValueType(value);
  if (isFormulaValue(value)) {
    const formula = stripLeadingEquals(value.formula);
    return formula === value.formula ? value : {...value, formula};
  }
  // A shared formula's `sharedFormula` is a master cell address, not formula text, so there is no
  // leading `=` to canonicalise — it passes through as given.
  return value;
}

/**
 * OOXML stores a formula's text in `<f>` without the UI's leading `=`. Strip a single
 * leading `=` so the stored form is canonical and a strict consumer (Google Sheets,
 * WPS) — which rejects a stored formula beginning with `=` — accepts the file.
 */
function stripLeadingEquals(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula;
}
