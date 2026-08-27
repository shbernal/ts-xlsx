// Address decoding: the foundational spreadsheet primitive.
//
// Every higher layer (cells, ranges, defined names, formulas) is ultimately
// addressed by a `col`/`row` pair, so this module is the first thing the rewrite
// builds and the bedrock everything else stands on. It is pure, allocation-bounded,
// and has no I/O or platform dependency.
//
// The honesty rule that drives the shape here: an axis that a reference does not
// mention is `undefined`, never a sentinel. A whole-row reference (`$1`) has no
// column; a whole-column reference (`$A:$A`) has no row. Legacy code let those
// absent axes decay into `NaN`/`"undefined"` and leak into serialized addresses
// (`"$undefined$1"`, `"NaN:NaN"`), the exact defect the corpus locks against.
//
// The decoders come in two arities and two temperaments. `decodeAddress`/`decodeCellRef`/
// `decodeRange` face a caller and throw, because a reference a caller supplied is a mistake at the
// call. `tryDecodeCellRef`/`tryDecodeRange` face a file and return `undefined`, because a file this
// library did not write is allowed to be wrong and losing one attribute beats losing the sheet.
// Read-side code uses the second pair and nothing else; the rule and the reasoning are `xml-read.ts`'s.

/** Excel's column bounds: `A` (1) through `XFD` (16384). */
export const MAX_COLUMN = 16384;

/** Excel's row bound: 1 through 1048576. The other axis of {@link MAX_COLUMN}. */
export const MAX_ROW = 1048576;

// The numeric door into the grid. `numberToColumn`/`columnToNumber` bound a reference spelled in
// letters; these bound the same position spelled as a number, so `getColumn(16385)` and
// `getCell('XFE1')` refuse the same mistake with the same words. Native `RangeError` rather than
// `AuthoringError`: a single scalar out of range is what `errors.ts` reserves for native errors.

/** @throws {RangeError} unless `n` is an integer in `1..MAX_COLUMN`. */
export function assertColumnInBounds(n: number): void {
  if (!Number.isInteger(n) || n < 1 || n > MAX_COLUMN) {
    throw new RangeError(`column ${n} is out of bounds: Excel supports 1..${MAX_COLUMN}`);
  }
}

/** @throws {RangeError} unless `n` is an integer in `1..MAX_ROW`. */
export function assertRowInBounds(n: number): void {
  if (!Number.isInteger(n) || n < 1 || n > MAX_ROW) {
    throw new RangeError(`row ${n} is out of bounds: Excel supports 1..${MAX_ROW}`);
  }
}

/**
 * A rectangular block of the grid, as **inclusive** 1-based bounds on both axes.
 *
 * One declaration because inclusive-first/last is the convention every range-shaped thing in this
 * library follows, and three copies of a convention are three places it can drift. A merged region,
 * a table's extent and a {@link Range} handle are all this shape; what differs between them is what
 * the rectangle *means*, which is what their own names carry.
 */
export interface GridRect {
  /** 1-based row of the top edge. */
  readonly top: number;
  /** 1-based column of the left edge. */
  readonly left: number;
  /** 1-based row of the bottom edge, inclusive. */
  readonly bottom: number;
  /** 1-based column of the right edge, inclusive. */
  readonly right: number;
}

/** Whether two grid rectangles share at least one cell. */
export function rectsOverlap(a: GridRect, b: GridRect): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** A decoded single-cell reference. An axis the reference omits is `undefined`. */
export interface CellAddress {
  /** Canonical A1 form with `$` anchors stripped: e.g. `"B2"`, `"1"`, `"A"`. */
  readonly address: string;
  /** 1-based column, or `undefined` for a row-only reference (`$1`). */
  readonly col: number | undefined;
  /** 1-based row, or `undefined` for a column-only reference (`$A`). */
  readonly row: number | undefined;
}

/**
 * A decoded range reference. Corners are the min/max of the endpoints per axis;
 * an axis neither endpoint mentions (a whole-row or whole-column range) is
 * `undefined` on every corner and simply absent from `dimensions`.
 */
export interface RangeAddress {
  readonly top: number | undefined;
  readonly left: number | undefined;
  readonly bottom: number | undefined;
  readonly right: number | undefined;
  /** The originating sheet, present only when the reference carried one. */
  readonly sheetName?: string;
  readonly tl: CellAddress;
  readonly br: CellAddress;
  /** Canonical `tl:br` form: `"A1:B2"`, `"1:1"` (rows), `"A:A"` (columns). */
  readonly dimensions: string;
}

const SINGLE_REF = /^\$?([A-Z]*)\$?(\d*)$/;
// A leading `Sheet!` prefix: quoted (`'a''b'!`, doubled apostrophe escapes one) or
// bare (`Sheet1!`). Group 1 = quoted body, group 2 = bare name, group 3 = the rest.
const SHEET_PREFIX = /^(?:(?:'((?:[^']|'')*)')|([^'!]+))!(.*)$/;

/** Convert a 1-based column number to its letters (`1 → "A"`, `27 → "AA"`). */
export function numberToColumn(n: number): string {
  if (!Number.isInteger(n) || n < 1 || n > MAX_COLUMN) {
    throw new RangeError(`column ${n} is out of bounds: Excel supports 1..${MAX_COLUMN}`);
  }
  let letters = '';
  let remaining = n;
  while (remaining > 0) {
    const digit = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + digit) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

/** Convert column letters to a 1-based number (`"A" → 1`, `"AA" → 27`). */
export function columnToNumber(letters: string): number {
  if (letters.length === 0 || letters.length > 3) {
    throw new RangeError(`invalid column letters: "${letters}"`);
  }
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i);
    if (code < 65 || code > 90) {
      throw new RangeError(`invalid column letters: "${letters}"`);
    }
    n = n * 26 + (code - 64);
  }
  if (n > MAX_COLUMN) {
    throw new RangeError(
      `column "${letters}" is out of bounds: Excel supports up to ${MAX_COLUMN} (XFD)`,
    );
  }
  return n;
}

/** Build a {@link CellAddress} corner straight from optional numeric axes: the address string is
 * assembled from the parts we already hold, so no encode-then-decode round-trip is needed. An axis the
 * corner omits stays `undefined`; both absent yields the empty address (`""`). */
function makeCellAddress(col: number | undefined, row: number | undefined): CellAddress {
  return {
    address: `${col !== undefined ? numberToColumn(col) : ''}${row !== undefined ? row : ''}`,
    col,
    row,
  };
}

/**
 * Decode a single cell/row/column reference into `{address, col, row}`. Anchoring
 * `$` signs are accepted and dropped; an absent axis is `undefined`.
 *
 * @throws {SyntaxError} if the reference mentions neither a column nor a row.
 */
export function decodeAddress(reference: string): CellAddress {
  const match = SINGLE_REF.exec(reference);
  if (!match) {
    throw new SyntaxError(`invalid cell reference: "${reference}"`);
  }
  const letters = match[1] ?? '';
  const digits = match[2] ?? '';
  if (letters.length === 0 && digits.length === 0) {
    throw new SyntaxError(`invalid cell reference: "${reference}"`);
  }
  const col = letters.length > 0 ? columnToNumber(letters) : undefined;
  const row = digits.length > 0 ? Number.parseInt(digits, 10) : undefined;
  return {address: `${letters}${digits}`, col, row};
}

/**
 * A reference that names one cell, both axes present. The narrowing of {@link CellAddress} that
 * most callers actually want: `decodeAddress` is deliberately three-shaped because a bare row
 * (`$1`) and a bare column (`$A`) are legitimate references, but a cell is where a value lives, and
 * every caller that needs one was re-deriving that invariant by hand.
 */
export interface CellPosition {
  readonly col: number;
  readonly row: number;
}

/**
 * Decode a reference that must name a single cell. Anchoring `$` signs are accepted and dropped.
 *
 * @throws {SyntaxError} if the reference is unparseable, or parses but omits an axis (`"A"`, `"1"`).
 */
export function decodeCellRef(reference: string): CellPosition {
  const {col, row} = decodeAddress(reference);
  if (col === undefined || row === undefined) {
    throw new SyntaxError(
      `"${reference}" is not a single-cell reference: it omits a column or row`,
    );
  }
  return {col, row};
}

// The tolerant half of the decoders, and the read side's only door to them.
//
// `xml-read.ts` states the rule these honour: a file the library did not write is allowed to be
// wrong, and losing one attribute beats losing the sheet. Every other foreign scalar already has
// its tolerant reader (`numInteger`, `boolTristate`, `enumToken`); a reference had none, so a
// malformed `r`/`ref`/`sqref` in an untrusted package aborted the whole read with a native
// `RangeError` or `SyntaxError`, outside the `XlsxError` taxonomy one `catch` clause answers.
//
// "Tolerant" here means *a reference naming something that can exist*, not merely one that parses:
// `A0` and `A1048577` parse fine and then blow up at the grid, which is the same failure one step
// later. Bounds are part of the question, so they are part of the answer.
//
// The authoring decoders above keep throwing. `getCell('A0')` from a caller is a mistake at the
// call, and the taxonomy deliberately reserves a native error for a single scalar out of range.

/** Whether a row number a reference produced can name a line of the grid. Columns need no
 * companion: `columnToNumber` already refuses letters past `XFD`, so a decoded column is in
 * bounds or the decode threw. */
function rowCanExist(row: number | undefined): boolean {
  return row === undefined || (Number.isInteger(row) && row >= 1 && row <= MAX_ROW);
}

/**
 * {@link decodeCellRef} for a reference that came out of a file rather than out of a caller:
 * `undefined` for anything that does not name one cell that can exist: a range, a bare row or
 * column, a position off the grid (`A0`, `XFE1`), or outright garbage. A foreign producer writes
 * all of them, and none is worth throwing over when the reading code's answer is simply "then
 * there is nothing here".
 */
export function tryDecodeCellRef(reference: string): CellPosition | undefined {
  let position: CellPosition;
  try {
    position = decodeCellRef(reference);
  } catch {
    return undefined;
  }
  return rowCanExist(position.row) ? position : undefined;
}

/**
 * {@link decodeRange} for a reference that came out of a file: `undefined` for anything that does
 * not name a region that can exist. The sibling of {@link tryDecodeCellRef} on the other arity: a
 * `ref` or one area of a `sqref` is as likely to be malformed as a cell's `r`, and the reader's
 * answer to both is the same one.
 *
 * An axis neither endpoint mentions stays `undefined`, exactly as in `decodeRange`: a whole-column
 * range is unbounded, not unreadable, and a caller that needs a bounded rectangle says so itself.
 */
export function tryDecodeRange(reference: string): RangeAddress | undefined {
  let range: RangeAddress;
  try {
    range = decodeRange(reference);
  } catch {
    return undefined;
  }
  return rowCanExist(range.top) && rowCanExist(range.bottom) ? range : undefined;
}

/**
 * Decode a range reference (`A1:B2`, `$1:$1`, `Sheet1!$A:$A`) into its corners and
 * canonical dimensions. A single reference collapses to a degenerate range whose
 * corners coincide.
 */
export function decodeRange(reference: string): RangeAddress {
  const prefix = SHEET_PREFIX.exec(reference);
  let sheetName: string | undefined;
  let body = reference;
  if (prefix) {
    const quoted = prefix[1];
    sheetName = quoted !== undefined ? quoted.replace(/''/g, "'") : prefix[2];
    body = prefix[3] ?? '';
  }

  const parts = body.split(':');
  const start = decodeAddress(parts[0] ?? '');
  const end = parts.length > 1 ? decodeAddress(parts[1] ?? '') : start;

  const cols = [start.col, end.col].filter((v): v is number => v !== undefined);
  const rows = [start.row, end.row].filter((v): v is number => v !== undefined);
  const left = cols.length > 0 ? Math.min(...cols) : undefined;
  const right = cols.length > 0 ? Math.max(...cols) : undefined;
  const top = rows.length > 0 ? Math.min(...rows) : undefined;
  const bottom = rows.length > 0 ? Math.max(...rows) : undefined;

  const tl = makeCellAddress(left, top);
  const br = makeCellAddress(right, bottom);

  return {
    top,
    left,
    bottom,
    right,
    ...(sheetName !== undefined ? {sheetName} : {}),
    tl,
    br,
    dimensions: `${tl.address}:${br.address}`,
  };
}

/** Encode a 1-based `col`/`row` pair into its canonical A1 address (`"B2"`). */
export function encodeAddress(col: number, row: number): string {
  if (!Number.isInteger(row) || row < 1) {
    throw new RangeError(`row ${row} is out of bounds: rows start at 1`);
  }
  return `${numberToColumn(col)}${row}`;
}
