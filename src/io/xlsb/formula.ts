// Decoding a BIFF12 `Ptg` token stream back into formula text — the one place where the binary and
// XML serialisations of a workbook are genuinely different *languages* rather than different spellings.
//
// An `.xlsx` stores `SUM(A1:A5)/COUNT(A1:A5)` as those nineteen characters. An `.xlsb` stores the same
// formula as a postfix (reverse-Polish) token stream: two range operands, two calls, a divide. So the
// decoder is a stack machine — each operand pushes its own text, each operator pops what it needs and
// pushes the joined result, and a well-formed stream leaves exactly one string behind.
//
// Two things make the reconstruction exact rather than approximate:
//
//   - **Parentheses are stored, not inferred.** Excel emits an explicit `PtgParen` wherever the author
//     typed one, so there is no precedence arithmetic here and no risk of `(1+2)*3` reading back as
//     `1+2*3`. The token stream already says which is which.
//   - **A reference names no sheet.** A 3-D reference carries an *index* into the workbook's
//     `BrtExternSheet` table, which in turn names a span of sheets in a supporting book. Resolving
//     that indirection — and re-quoting the sheet name the way Excel would — is what turns token
//     `ixti=2` back into `'Odd Name'!A1`.
//
// **A token this decoder does not know makes the whole formula undecodable, by design.** The stream is
// self-describing only if every token's length is known, so guessing past an unrecognised token would
// desynchronise the parse and produce confident nonsense. Instead the decoder returns `undefined` and
// its caller keeps what it can still trust — the cached result Excel stored beside the formula. The
// gaps that reach that path are listed in `docs/knowledge/specs/xlsb-binary-format-output.md`.

import {MAX_COLUMN, numberToColumn} from '../../core/address.ts';
import {quoteSheetName} from '../../core/formula.ts';
import {errorCodeFor, RecordReader} from './primitives.ts';
import {FTAB_USER_DEFINED, fixedArityFor, functionNameFor} from './ptg-functions.ts';

/** One `Xti` ([MS-XLSB] 2.5.163): the span of sheets, in one supporting book, that an `ixti` names. */
export interface ExternSheetRef {
  readonly supBook: number;
  /** Zero-based sheet index; the span is inclusive. */
  readonly firstSheet: number;
  readonly lastSheet: number;
}

/** The workbook-level tables a formula's references and names resolve through. */
export interface FormulaScope {
  /** Sheet names in workbook (tab) order — what an `Xti`'s sheet indices point into. */
  readonly sheetNames: readonly string[];
  /** The `BrtExternSheet` table, indexed by a 3-D token's `ixti`. */
  readonly externSheets: readonly ExternSheetRef[];
  /** The index of the supporting book that is this workbook, or `undefined` when the file declares a
   * supporting book this reader does not recognise. A workbook with no external links declares
   * exactly one — itself — so this is the ordinary case; anything else leaves the indices untrustworthy
   * and no 3-D reference resolves, which drops those formulas rather than naming the wrong sheet. */
  readonly selfSupBook: number | undefined;
  /** Every `BrtName` in file order, function placeholders included — a `PtgName` cites one by
   * **1-based** index, so filtering this list would misaddress every name reference. */
  readonly names: readonly string[];
}

/** The cell a `PtgExp`/`PtgTbl` stream defers to: the top-left of the shared or array formula group. */
export interface FormulaAnchor {
  /** Zero-based. */
  readonly row: number;
  readonly column: number;
}

/**
 * Decode a `CellParsedFormula`'s token stream into formula text, in the same on-disk spelling the XML
 * form writes into `<f>` — `_xlfn.`-prefixed function names included, so the caller applies the same
 * `unmangleFunctions` normalisation to either serialisation.
 *
 * @param rgce the token stream.
 * @param rgcb the trailing extra-data block: the array constants, and the cell ranges a precomputed
 *   range token refers to. Its entries are consumed in token order.
 * @returns the formula text, or `undefined` if the stream uses a token this reader does not decode.
 * @throws {XlsbParseError} if a token runs past the end of the stream (a malformed formula).
 */
export function decodeFormula(
  rgce: Uint8Array,
  rgcb: Uint8Array,
  scope: FormulaScope,
): string | undefined {
  const tokens = new RecordReader(rgce);
  const extra = new RecordReader(rgcb);
  const stack: string[] = [];
  // Pushing `undefined` is how an undecodable token is reported without unwinding: the loop stops and
  // the arity check below rejects the stream. It keeps every token case a plain expression.
  const push = (text: string | undefined): boolean => {
    if (text === undefined) return false;
    stack.push(text);
    return true;
  };

  while (!tokens.done) {
    if (!step(tokens.u8(), tokens, extra, stack, scope, push)) return undefined;
  }
  return stack.length === 1 ? stack[0] : undefined;
}

/**
 * The master cell a token stream defers to, when the stream is nothing but a `PtgExp` (a shared or
 * array formula's member) or a `PtgTbl` (a data-table cell) rather than a formula of its own.
 *
 * The column lives in the extra-data block rather than the token, which is the one place BIFF12
 * splits a single reference across the two halves of a `CellParsedFormula`.
 */
export function formulaAnchor(rgce: Uint8Array, rgcb: Uint8Array): FormulaAnchor | undefined {
  if (rgce.length !== 5 || (rgce[0] !== PTG.Exp && rgce[0] !== PTG.Tbl)) return undefined;
  const tokens = new RecordReader(rgce);
  tokens.skip(1);
  const row = tokens.u32();
  const extra = new RecordReader(rgcb);
  if (extra.remaining < 4) return undefined;
  return {row, column: extra.u32()};
}

// One token: decode it, mutate the stack, and report whether the stream is still decodable.
function step(
  ptg: number,
  tokens: RecordReader,
  extra: RecordReader,
  stack: string[],
  scope: FormulaScope,
  push: (text: string | undefined) => boolean,
): boolean {
  const binary = BINARY_OPERATORS.get(ptg);
  if (binary !== undefined) {
    const right = stack.pop();
    const left = stack.pop();
    return left !== undefined && right !== undefined && push(`${left}${binary}${right}`);
  }

  switch (ptg) {
    case PTG.Uplus:
    case PTG.Uminus:
    case PTG.Percent: {
      const operand = stack.pop();
      if (operand === undefined) return false;
      return push(
        ptg === PTG.Percent ? `${operand}%` : `${ptg === PTG.Uplus ? '+' : '-'}${operand}`,
      );
    }
    case PTG.Paren: {
      const operand = stack.pop();
      return operand !== undefined && push(`(${operand})`);
    }
    case PTG.MissArg:
      // An omitted argument — `IF(A1>0,,1)` — is a real operand whose text is nothing at all.
      return push('');
    case PTG.Str:
      return push(quoteString(tokens.shortString()));
    case PTG.Attr:
      return attribute(tokens, stack, push);
    case PTG.Err:
      return push(errorCodeFor(tokens.u8()));
    case PTG.Bool:
      return push(tokens.u8() !== 0 ? 'TRUE' : 'FALSE');
    case PTG.Int:
      return push(String(tokens.u16()));
    case PTG.Num:
      return push(numberText(tokens.f64()));
    default:
      // Every remaining token is an operand or call whose meaning is independent of its result class
      // (reference, value, or array) — the class only tells the calculation engine how to coerce it.
      return ptg >= CLASSED_TOKEN_FLOOR
        ? operand(
            (ptg & CLASSED_TOKEN_MASK) | CLASSED_TOKEN_FLOOR,
            tokens,
            extra,
            stack,
            scope,
            push,
          )
        : false;
  }
}

// A class-carrying operand or call token, reduced to its base ptg.
function operand(
  base: number,
  tokens: RecordReader,
  extra: RecordReader,
  stack: string[],
  scope: FormulaScope,
  push: (text: string | undefined) => boolean,
): boolean {
  switch (base) {
    case PTG.Array:
      tokens.skip(14); // A size hint the extra-data block restates; the block is the authority.
      return push(arrayConstant(extra));
    case PTG.Func: {
      const name = functionNameFor(tokens.u16());
      const arity = name === undefined ? undefined : fixedArityFor(name);
      return name !== undefined && arity !== undefined && push(call(name, arity, stack));
    }
    case PTG.FuncVar:
      return variadicCall(tokens, stack, push);
    case PTG.Name: {
      // Cited 1-based, and into the *unfiltered* name list — the placeholder names Excel registers for
      // post-2007 functions occupy indices too, even though they are not the workbook's defined names.
      return push(scope.names[tokens.u32() - 1]);
    }
    case PTG.Ref:
      return push(cellText(tokens.u32(), tokens.u16()));
    case PTG.Area:
      return push(rangeText(tokens.u32(), tokens.u32(), tokens.u16(), tokens.u16()));
    case PTG.Ref3d: {
      const prefix = sheetPrefix(tokens.u16(), scope);
      const text = cellText(tokens.u32(), tokens.u16());
      return prefix !== undefined && text !== undefined && push(prefix + text);
    }
    case PTG.Area3d: {
      const prefix = sheetPrefix(tokens.u16(), scope);
      const text = rangeText(tokens.u32(), tokens.u32(), tokens.u16(), tokens.u16());
      return prefix !== undefined && text !== undefined && push(prefix + text);
    }
    case PTG.RefErr:
      tokens.skip(6);
      return push(REFERENCE_ERROR);
    case PTG.AreaErr:
      tokens.skip(12);
      return push(REFERENCE_ERROR);
    case PTG.RefErr3d:
      tokens.skip(8);
      return push(REFERENCE_ERROR);
    case PTG.AreaErr3d:
      tokens.skip(14);
      return push(REFERENCE_ERROR);
    case PTG.MemArea:
      // A precomputed range: the tokens it was computed from follow inline, so the header is skipped
      // and the walk simply continues into them. Its extra-data entry — the resulting rectangles — is
      // a calculation shortcut with nothing to say about the text, but must still be consumed in order.
      tokens.skip(6);
      return skipExtraRanges(extra);
    default:
      return false;
  }
}

// `PtgAttr` ([MS-XLSB] 2.5.97.1): a family of hints the calculation engine leaves in the stream.
// Almost all are invisible in the formula text — the jump offsets an `IF` uses to skip the branch it
// did not take, the marker on a volatile function. The one that carries meaning is `bitSum`, Excel's
// encoding of a single-argument `SUM`, which is a call by any other name.
function attribute(
  tokens: RecordReader,
  stack: string[],
  push: (text: string | undefined) => boolean,
): boolean {
  const flags = tokens.u8();
  const data = tokens.u16();
  if ((flags & ATTR_CHOOSE) !== 0) {
    // The one variable-length attribute: `data` counts CHOOSE's branches, each with a jump offset,
    // plus one for the end of the call.
    tokens.skip((data + 1) * 2);
    return true;
  }
  if ((flags & ATTR_SUM) !== 0) return push(call('SUM', 1, stack));
  // `bitSpace` records whitespace the author typed around a token. It is cosmetic — Excel redisplays
  // the formula identically without it — and reattaching it to the right operand is not something a
  // postfix walk can do, so it is dropped rather than misplaced.
  return true;
}

// `PtgFuncVar` ([MS-XLSB] 2.5.97.4): a call whose argument count is in the token. Index 255 is not a
// function at all but the indirection every post-2007 function is called through: the name comes from
// the stream's first operand, which is a `PtgName` pointing at Excel's `_xlfn.`-prefixed placeholder.
function variadicCall(
  tokens: RecordReader,
  stack: string[],
  push: (text: string | undefined) => boolean,
): boolean {
  const count = tokens.u8() & FUNCVAR_PARAM_MASK;
  const index = tokens.u16() & FUNCVAR_INDEX_MASK;
  if (index !== FTAB_USER_DEFINED) {
    const name = functionNameFor(index);
    return name !== undefined && push(call(name, count, stack));
  }
  if (count < 1 || stack.length < count) return false;
  const parts = stack.splice(stack.length - count, count);
  const [name, ...args] = parts as [string, ...string[]];
  return push(`${name}(${args.join(',')})`);
}

// Pop `arity` arguments and push the call. Arguments were pushed left to right, so they come off the
// stack as one contiguous run in source order.
function call(name: string, arity: number, stack: string[]): string | undefined {
  if (stack.length < arity) return undefined;
  return `${name}(${stack.splice(stack.length - arity, arity).join(',')})`;
}

// `PtgExtraArray` ([MS-XLSB] 2.5.97.2): the elements of an array constant, row-major, behind a
// row/column count. The element encodings are fixed-width apart from the string, which carries its
// own length — so the block is walked, never indexed.
function arrayConstant(extra: RecordReader): string | undefined {
  const rows = extra.u32();
  const columns = extra.u32();
  if (rows === 0 || columns === 0 || rows * columns > MAX_ARRAY_ELEMENTS) return undefined;
  const lines: string[] = [];
  for (let row = 0; row < rows; row++) {
    const cells: string[] = [];
    for (let column = 0; column < columns; column++) {
      const element = arrayElement(extra);
      if (element === undefined) return undefined;
      cells.push(element);
    }
    lines.push(cells.join(','));
  }
  return `{${lines.join(';')}}`;
}

// One `SerAr` ([MS-XLSB] 2.5.129) element of an array constant.
function arrayElement(extra: RecordReader): string | undefined {
  switch (extra.u8()) {
    case SER_NUM:
      return numberText(extra.f64());
    case SER_STR:
      return quoteString(extra.shortString());
    case SER_BOOL:
      return extra.u8() !== 0 ? 'TRUE' : 'FALSE';
    case SER_ERR: {
      const text = errorCodeFor(extra.u8());
      extra.skip(3);
      return text;
    }
    default:
      return undefined;
  }
}

// Consume the `PtgExtraMem` a precomputed-range token owns: a count, then that many cell ranges.
function skipExtraRanges(extra: RecordReader): boolean {
  const count = extra.u32();
  if (count * RANGE_BYTES > extra.remaining) return false;
  extra.skip(count * RANGE_BYTES);
  return true;
}

// The sheet part of a 3-D reference, `Data!` or `'Odd Name'!` or `Data:More!`, from the index the
// token carries into the workbook's extern-sheet table.
function sheetPrefix(ixti: number, scope: FormulaScope): string | undefined {
  const xti = scope.externSheets[ixti];
  if (xti === undefined || xti.supBook !== scope.selfSupBook) return undefined;
  const first = scope.sheetNames[xti.firstSheet];
  const last = scope.sheetNames[xti.lastSheet];
  if (first === undefined || last === undefined) return undefined;
  return `${first === last ? quoteSheetName(first) : quoteSheetName(first, last)}!`;
}

// A single cell, from the row and the packed column word every reference token shares. The two high
// bits of that word say whether each axis is relative, which is exactly where the `$` signs go.
function cellText(row: number, packedColumn: number): string | undefined {
  const column = packedColumn & COLUMN_MASK;
  if (row > MAX_ROW_INDEX || column >= MAX_COLUMN) return undefined;
  const columnAbs = (packedColumn & COLUMN_RELATIVE) === 0 ? '$' : '';
  const rowAbs = (packedColumn & ROW_RELATIVE) === 0 ? '$' : '';
  return `${columnAbs}${numberToColumn(column + 1)}${rowAbs}${row + 1}`;
}

// A range. A range that spans every row of its columns, or every column of its rows, is written in
// Excel's abbreviated form (`A:A`, `2:2`) — which is not cosmetic: it is the only spelling Excel
// writes for a whole-column reference, so anything else would fail to match the XML twin.
function rangeText(
  rowFirst: number,
  rowLast: number,
  packedFirst: number,
  packedLast: number,
): string | undefined {
  if (rowFirst === 0 && rowLast === MAX_ROW_INDEX) {
    const first = columnOnly(packedFirst);
    const last = columnOnly(packedLast);
    return first === undefined || last === undefined ? undefined : `${first}:${last}`;
  }
  if ((packedFirst & COLUMN_MASK) === 0 && (packedLast & COLUMN_MASK) === MAX_COLUMN - 1) {
    return rowFirst > MAX_ROW_INDEX || rowLast > MAX_ROW_INDEX
      ? undefined
      : `${rowOnly(rowFirst, packedFirst)}:${rowOnly(rowLast, packedLast)}`;
  }
  const first = cellText(rowFirst, packedFirst);
  const last = cellText(rowLast, packedLast);
  return first === undefined || last === undefined ? undefined : `${first}:${last}`;
}

function columnOnly(packedColumn: number): string | undefined {
  const column = packedColumn & COLUMN_MASK;
  if (column >= MAX_COLUMN) return undefined;
  return `${(packedColumn & COLUMN_RELATIVE) === 0 ? '$' : ''}${numberToColumn(column + 1)}`;
}

function rowOnly(row: number, packedColumn: number): string {
  return `${(packedColumn & ROW_RELATIVE) === 0 ? '$' : ''}${row + 1}`;
}

// A string literal, in the formula's own escaping: the delimiter is a double quote, and a double quote
// inside the text is doubled.
function quoteString(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

// A numeric literal. JavaScript and Excel agree on every ordinary number; they part company only at
// the exponent's case, which is normalised here so `1E+21` does not read back as `1e+21`.
function numberText(value: number): string {
  return String(value).toUpperCase();
}

const REFERENCE_ERROR = '#REF!';

// The infix operators, by their ptg. `PtgIsect` is Excel's space operator (`A1:A3 A2:A5`) and
// `PtgUnion` its comma — both are operators despite looking like punctuation.
const BINARY_OPERATORS: ReadonlyMap<number, string> = new Map([
  [0x03, '+'],
  [0x04, '-'],
  [0x05, '*'],
  [0x06, '/'],
  [0x07, '^'],
  [0x08, '&'],
  [0x09, '<'],
  [0x0a, '<='],
  [0x0b, '='],
  [0x0c, '>='],
  [0x0d, '>'],
  [0x0e, '<>'],
  [0x0f, ' '],
  [0x10, ','],
  [0x11, ':'],
]);

// The token numbers this decoder names. Operand tokens are listed at their *base* value: the stream
// carries them offset by a result class (+0x20 value, +0x40 array), which changes how the calculation
// engine coerces the operand but never what it says.
const PTG = {
  Exp: 0x01,
  Tbl: 0x02,
  Uplus: 0x12,
  Uminus: 0x13,
  Percent: 0x14,
  Paren: 0x15,
  MissArg: 0x16,
  Str: 0x17,
  Attr: 0x19,
  Err: 0x1c,
  Bool: 0x1d,
  Int: 0x1e,
  Num: 0x1f,
  Array: 0x20,
  Func: 0x21,
  FuncVar: 0x22,
  Name: 0x23,
  Ref: 0x24,
  Area: 0x25,
  MemArea: 0x26,
  RefErr: 0x2a,
  AreaErr: 0x2b,
  Ref3d: 0x3a,
  Area3d: 0x3b,
  RefErr3d: 0x3c,
  AreaErr3d: 0x3d,
} as const;

const CLASSED_TOKEN_FLOOR = 0x20;
const CLASSED_TOKEN_MASK = 0x1f;

// `PtgAttr` flag bits.
const ATTR_CHOOSE = 0x04;
const ATTR_SUM = 0x10;

// `PtgFuncVar`'s two packed fields each reserve their top bit for a flag the text does not carry.
const FUNCVAR_PARAM_MASK = 0x7f;
const FUNCVAR_INDEX_MASK = 0x7fff;

// The packed column word shared by every reference token: 14 bits of column, then the two relative-
// axis flags.
const COLUMN_MASK = 0x3fff;
const COLUMN_RELATIVE = 0x4000;
const ROW_RELATIVE = 0x8000;
const MAX_ROW_INDEX = 1048575;

// `SerAr` element tags.
const SER_NUM = 0x00;
const SER_STR = 0x01;
const SER_BOOL = 0x02;
const SER_ERR = 0x04;

// An `UncheckedRfX` is four 32-bit bounds.
const RANGE_BYTES = 16;

// A bound on an array constant's declared size. The elements themselves are read from the extra-data
// block, which cannot outrun its own record — but the row × column product is multiplied *before* any
// of it is read, and a forged pair would otherwise buy a loop of its own choosing. Excel's own limit
// on an array constant is far below this.
const MAX_ARRAY_ELEMENTS = 1 << 20;
