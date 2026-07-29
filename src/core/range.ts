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
import {applyCellStyle, type Cell} from './cell.ts';
import {
  type Alignment,
  type Border,
  CELL_STYLE_FACETS,
  type CellStyle,
  type Fill,
  type Font,
  type Protection,
} from './style.ts';
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

  /**
   * The block's style, facet by facet — the counterpart of {@link Cell.style} over a rectangle, with
   * the same semantics in both directions.
   *
   * **Reading** reports a facet only when *every* position in the block carries a structurally
   * identical one, and `undefined` when they differ or any position is still empty. A block styled
   * through this handle therefore reads back what was written; a block whose cells disagree says so
   * rather than picking a corner's answer and passing it off as the whole.
   *
   * **Writing** lays each facet the payload names onto every cell, leaving facets it omits untouched
   * — exactly what `cell.style = {...}` does, so this composes with prior styling instead of clearing
   * it. Use {@link clearStyle} first for a wholesale replace.
   *
   * Writing **materialises** every position in the block, because a styled-but-valueless cell is the
   * only way an empty cell renders with a fill: skipping the holes would leave gaps in a header band.
   * The cost is bounded by construction — a range is always a bounded rectangle, and whole-axis
   * styling belongs to {@link Worksheet.getColumn}/{@link Worksheet.getRow} instead. {@link cellCount}
   * is the exact number of cells a write will create.
   */
  get style(): CellStyle {
    const style: CellStyle = {};
    for (const facet of CELL_STYLE_FACETS) this.#collectShared(style, facet);
    return style;
  }

  set style(style: Readonly<CellStyle>) {
    for (const cell of this.#materialise()) applyCellStyle(cell, style);
  }

  /**
   * Strip every style facet from every cell in the block, leaving values untouched. Assigning
   * {@link style} composes, so this is how a wholesale replace is said: `clearStyle()` then assign.
   * Materialises nothing — a cell that does not exist carries no style to clear.
   */
  clearStyle(): void {
    for (const cell of this.cells) {
      for (const facet of CELL_STYLE_FACETS) setFacet(cell, facet, undefined);
    }
  }

  /** Fill applied to every cell in the block; `undefined` when they do not all agree. */
  get fill(): Fill | undefined {
    return this.#sharedFacet('fill');
  }
  set fill(fill: Fill | undefined) {
    this.#writeFacet('fill', fill);
  }

  /** Number format applied to every cell in the block; `undefined` when they do not all agree. */
  get numFmt(): string | undefined {
    return this.#sharedFacet('numFmt');
  }
  set numFmt(numFmt: string | undefined) {
    this.#writeFacet('numFmt', numFmt);
  }

  /** Font applied to every cell in the block; `undefined` when they do not all agree. */
  get font(): Font | undefined {
    return this.#sharedFacet('font');
  }
  set font(font: Font | undefined) {
    this.#writeFacet('font', font);
  }

  /** Border applied to every cell in the block; `undefined` when they do not all agree. */
  get border(): Border | undefined {
    return this.#sharedFacet('border');
  }
  set border(border: Border | undefined) {
    this.#writeFacet('border', border);
  }

  /** Alignment applied to every cell in the block; `undefined` when they do not all agree. */
  get alignment(): Alignment | undefined {
    return this.#sharedFacet('alignment');
  }
  set alignment(alignment: Alignment | undefined) {
    this.#writeFacet('alignment', alignment);
  }

  /** Protection flags applied to every cell in the block; `undefined` when they do not all agree. */
  get protection(): Protection | undefined {
    return this.#sharedFacet('protection');
  }
  set protection(protection: Protection | undefined) {
    this.#writeFacet('protection', protection);
  }

  // Every cell in the block, created where it does not exist yet. An address covered by a merge
  // resolves to that region's master, so a block overlapping a merge restyles the master rather than
  // stranding a style on a covered cell the serializer would then have to drop.
  #materialise(): Cell[] {
    return [...this.addresses()].map((address) => this.#sheet.getCell(address));
  }

  // Assigning a facet replaces that facet on every cell — the block-wide reading of `cell.fill = x`.
  // Clearing one (`undefined`) touches only the cells that exist: there is nothing to clear on a hole,
  // and materialising the block to write nothing onto it would be pure cost.
  #writeFacet<K extends keyof CellStyle>(facet: K, value: CellStyle[K]): void {
    const cells = value === undefined ? this.cells : this.#materialise();
    for (const cell of cells) setFacet(cell, facet, value);
  }

  // One facet of the block-wide style onto the record being assembled. Narrowed to a single key for
  // the same reason {@link setFacet} is — see there.
  #collectShared<K extends keyof CellStyle>(target: CellStyle, facet: K): void {
    const shared = this.#sharedFacet(facet);
    if (shared !== undefined) target[facet] = shared;
  }

  // A facet's value when every position in the block carries a structurally identical one, else
  // undefined. A hole counts as "no facet", so a partly-styled block is reported as disagreeing —
  // which it does, since the empty positions render unstyled.
  #sharedFacet<K extends keyof CellStyle>(facet: K): CellStyle[K] {
    let first: CellStyle[K] | undefined;
    let firstKey: string | undefined;
    let seen = 0;
    for (let row = this.top; row <= this.bottom; row++) {
      for (let col = this.left; col <= this.right; col++) {
        const value = this.#sheet.hasCell(row, col)
          ? this.#sheet.getCell(encodeAddress(col, row))[facet]
          : undefined;
        const key = facetKey(value);
        if (seen === 0) {
          first = value;
          firstKey = key;
        } else if (key !== firstKey) {
          return undefined;
        }
        seen++;
      }
    }
    return first;
  }
}

// Write one facet through a `Cell`'s own setter. Narrowed to a single key — rather than assigning
// `cell[facet]` with `facet` still a union — for the same reason `style.ts`'s `copyFacet` is: the
// compiler cannot correlate a union key with its value type across two object types, and a cast here
// would be the one place a facet could be written to the wrong slot without anything noticing.
function setFacet<K extends keyof CellStyle>(cell: Cell, facet: K, value: CellStyle[K]): void {
  const target: CellStyle = cell;
  target[facet] = value;
}

// A canonical string for a facet value, so two structurally identical records compare equal whatever
// order their keys were written in. Sound here and nowhere near a general deep-equal: every facet is
// a plain data record of strings, numbers, booleans and nested records — no functions, no cycles, no
// class instances — which is exactly the shape JSON round-trips faithfully.
function facetKey(value: unknown): string {
  if (value === undefined) return ' ';
  return JSON.stringify(value, (_key, inner: unknown) =>
    inner !== null && typeof inner === 'object' && !Array.isArray(inner)
      ? Object.fromEntries(Object.entries(inner as object).sort(([a], [b]) => (a < b ? -1 : 1)))
      : inner,
  );
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
