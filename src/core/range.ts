// A handle on a rectangular block of cells, reached by an A1 range reference or by its corners.
//
// The third axis handle, after `Row` and `Column`, and the same contract as both: a live view over
// the worksheet's stores, bounds fixed at construction, nothing materialised on read. See
// `core/row.ts` for why these are handles rather than records.
//
// **The rectangle is bounded, always.** `A:A` and `1:1` are legal range references and are rejected
// here, because a whole-axis "range" is not this shape of thing: styling one means stamping a
// million cells, whereas OOXML expresses it natively as a single `<col style>` or `<row s>` — which
// is what `getColumn`/`getRow` already write, in constant space. Refusing them is what lets a range
// materialise its cells eagerly without a cost cliff hiding behind an innocuous-looking call.

import {decodeRange, encodeAddress, MAX_COLUMN, MAX_ROW} from './address.ts';
import type {Cell} from './cell.ts';
import type {Worksheet} from './worksheet.ts';

/**
 * A rectangular block of a worksheet's cells: `sheet.getRange('B2:D5')`.
 *
 * Cheap and stateless — constructing one creates no cells and does not extend the used range.
 * {@link addresses} walks the block without materialising anything; {@link cells} reports only what
 * already exists.
 *
 * Bounds are **inclusive first/last**, never start-and-count. That is the convention every
 * range-shaped accessor in this library follows, so the three axes cannot disagree about what a
 * pair of numbers means.
 */
export class Range {
  readonly #sheet: Worksheet;

  /** 1-based row of the block's top edge. Fixed for this handle's lifetime. */
  readonly top: number;
  /** 1-based column of the block's left edge. Fixed for this handle's lifetime. */
  readonly left: number;
  /** 1-based row of the block's bottom edge, inclusive. */
  readonly bottom: number;
  /** 1-based column of the block's right edge, inclusive. */
  readonly right: number;

  /**
   * Build a handle from inclusive corners, in any order — `(5, 4, 2, 2)` and `(2, 2, 5, 4)` name the
   * same block, exactly as `D5:B2` and `B2:D5` do. Prefer {@link Worksheet.getRange}.
   *
   * @throws {RangeError} if a corner is not a positive integer or falls outside the sheet's bounds.
   */
  constructor(sheet: Worksheet, top: number, left: number, bottom: number, right: number) {
    checkBound('row', top, MAX_ROW);
    checkBound('row', bottom, MAX_ROW);
    checkBound('column', left, MAX_COLUMN);
    checkBound('column', right, MAX_COLUMN);
    this.#sheet = sheet;
    this.top = Math.min(top, bottom);
    this.bottom = Math.max(top, bottom);
    this.left = Math.min(left, right);
    this.right = Math.max(left, right);
  }

  /** The worksheet this block belongs to. */
  get sheet(): Worksheet {
    return this.#sheet;
  }

  /** Canonical `tl:br` A1 form — `"B2:D5"`. A one-cell block still reads as `"B2:B2"`. */
  get address(): string {
    return `${encodeAddress(this.left, this.top)}:${encodeAddress(this.right, this.bottom)}`;
  }

  /** Rows spanned, inclusive of both edges. */
  get rowCount(): number {
    return this.bottom - this.top + 1;
  }

  /** Columns spanned, inclusive of both edges. */
  get columnCount(): number {
    return this.right - this.left + 1;
  }

  /** Cells the block covers — `rowCount * columnCount`, whether or not they exist yet. */
  get cellCount(): number {
    return this.rowCount * this.columnCount;
  }

  /** Whether a 1-based position falls inside the block. */
  contains(row: number, col: number): boolean {
    return row >= this.top && row <= this.bottom && col >= this.left && col <= this.right;
  }

  /**
   * Every address the block covers, row-major (`B2`, `C2`, `D2`, `B3`, …). Materialises nothing, so
   * this is the cheap way to walk a large block — and, being a generator, it can be abandoned
   * part-way without having built the whole list.
   */
  *addresses(): IterableIterator<string> {
    for (let row = this.top; row <= this.bottom; row++) {
      for (let col = this.left; col <= this.right; col++) {
        yield encodeAddress(col, row);
      }
    }
  }

  /**
   * The block's **materialised** cells, row-major. Sparse: a position nothing has ever written to is
   * simply absent, which is what distinguishes "never written" from a cell holding `null`. Reading
   * this creates nothing — mirroring {@link Column.cells}.
   */
  get cells(): readonly Cell[] {
    const cells: Cell[] = [];
    for (let row = this.top; row <= this.bottom; row++) {
      for (let col = this.left; col <= this.right; col++) {
        if (this.#sheet.hasCell(row, col)) cells.push(this.#sheet.getCell(encodeAddress(col, row)));
      }
    }
    return cells;
  }
}

function checkBound(axis: 'row' | 'column', value: number, max: number): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new RangeError(`${axis} ${value} is out of bounds — ${axis}s start at 1`);
  }
  if (value > max) {
    throw new RangeError(`${axis} ${value} is out of bounds — the sheet ends at ${max}`);
  }
}

/**
 * Resolve an A1 range reference against a worksheet into a {@link Range}.
 *
 * A reference carrying a sheet prefix is accepted only when it names *this* sheet (case-insensitively,
 * as sheet lookup is everywhere else): a range handed out by `sheetA` cannot be a block of `sheetB`,
 * and quietly ignoring the prefix would let a pasted `Sheet2!B2:D5` style the wrong sheet.
 *
 * @throws {SyntaxError} if the reference is unparseable, names another sheet, or leaves an axis
 *   unbounded (`A:A`, `1:1`).
 */
export function rangeFrom(sheet: Worksheet, reference: string): Range {
  const {top, left, bottom, right, sheetName, dimensions} = decodeRange(reference);
  if (sheetName !== undefined && sheetName.toLowerCase() !== sheet.name.toLowerCase()) {
    throw new SyntaxError(
      `"${reference}" names worksheet "${sheetName}", not "${sheet.name}" — a range belongs to the sheet it came from`,
    );
  }
  // An unbounded *row* axis is what `A:A` has: it names whole columns, every row of them.
  if (top === undefined || bottom === undefined) {
    throw new SyntaxError(
      `"${reference}" spans whole columns (${dimensions}) — style them through getColumn(n), which says the same thing in one attribute instead of ${MAX_ROW} cells`,
    );
  }
  if (left === undefined || right === undefined) {
    throw new SyntaxError(
      `"${reference}" spans whole rows (${dimensions}) — style them through getRow(n), which says the same thing in one attribute instead of ${MAX_COLUMN} cells`,
    );
  }
  return new Range(sheet, top, left, bottom, right);
}
