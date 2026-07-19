// Address decoding — the foundational spreadsheet primitive.
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
// (`"$undefined$1"`, `"NaN:NaN"`) — the exact defect the corpus locks against.

/** Excel's column bounds: `A` (1) through `XFD` (16384). */
export const MAX_COLUMN = 16384;

/** A decoded single-cell reference. An axis the reference omits is `undefined`. */
export interface CellAddress {
  /** Canonical A1 form with `$` anchors stripped — e.g. `"B2"`, `"1"`, `"A"`. */
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
  /** Canonical `tl:br` form — `"A1:B2"`, `"1:1"` (rows), `"A:A"` (columns). */
  readonly dimensions: string;
}

const SINGLE_REF = /^\$?([A-Z]*)\$?(\d*)$/;
// A leading `Sheet!` prefix: quoted (`'a''b'!`, doubled apostrophe escapes one) or
// bare (`Sheet1!`). Group 1 = quoted body, group 2 = bare name, group 3 = the rest.
const SHEET_PREFIX = /^(?:(?:'((?:[^']|'')*)')|([^'!]+))!(.*)$/;

/** Convert a 1-based column number to its letters (`1 → "A"`, `27 → "AA"`). */
export function numberToColumn(n: number): string {
  if (!Number.isInteger(n) || n < 1 || n > MAX_COLUMN) {
    throw new RangeError(`column ${n} is out of bounds — Excel supports 1..${MAX_COLUMN}`);
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
      `column "${letters}" is out of bounds — Excel supports up to ${MAX_COLUMN} (XFD)`,
    );
  }
  return n;
}

/** Build a {@link CellAddress} corner straight from optional numeric axes — the address string is
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
    throw new RangeError(`row ${row} is out of bounds — rows start at 1`);
  }
  return `${numberToColumn(col)}${row}`;
}
