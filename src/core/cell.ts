// A single cell: a typed value at a fixed 1-based `(row, col)` position.
//
// Position is immutable and numeric — the corpus locks `col`/`row` as 1-based
// *numbers* (legacy shipped a type declaration calling them strings, which broke
// strict consumers). The value is the only mutable state here; assigning it routes
// through the value model so the cell's `type` is always consistent with what it holds.

import {encodeAddress} from './address.ts';
import type {Fill, Font} from './style.ts';
import {type CellValue, type ValueType, coerceCellValue, detectValueType} from './value.ts';

export class Cell {
  /** 1-based row index. */
  readonly row: number;
  /** 1-based column index. */
  readonly col: number;

  #value: CellValue = null;
  #fill: Fill | undefined;
  #numFmt: string | undefined;
  #font: Partial<Font> | undefined;

  constructor(row: number, col: number) {
    if (!Number.isInteger(row) || row < 1) {
      throw new RangeError(`cell row ${row} is out of bounds — rows start at 1`);
    }
    if (!Number.isInteger(col) || col < 1) {
      throw new RangeError(`cell column ${col} is out of bounds — columns start at 1`);
    }
    this.row = row;
    this.col = col;
  }

  /** Canonical A1 address of this cell (`"B3"`). */
  get address(): string {
    return encodeAddress(this.col, this.row);
  }

  /** The cell's value; `null` when empty. Assigning `undefined` clears it. */
  get value(): CellValue {
    return this.#value;
  }

  set value(value: CellValue | undefined) {
    this.#value = coerceCellValue(value);
  }

  /** The observable {@link ValueType} of the current value. */
  get type(): ValueType {
    return detectValueType(this.#value);
  }

  /**
   * The cell's background fill, or `undefined` when it has none. Each cell owns its
   * own fill; assigning one never aliases a neighbour's style, so a fill set on one
   * cell cannot bleed onto its row, column, or sheet siblings.
   */
  get fill(): Fill | undefined {
    return this.#fill;
  }

  set fill(fill: Fill | undefined) {
    this.#fill = fill;
  }

  /**
   * The cell's number-format code (`"0.00%"`, a custom accounting format, …), or
   * `undefined` for the General format. Stored verbatim: the invariant form Excel
   * persists — `.` decimal, `,` grouping, `/` date separator — is neither localized
   * nor rewritten, so the code round-trips character-for-character. Like {@link fill},
   * each cell owns its own code; a cell that also carries a column-level format keeps
   * both, so overriding one facet never drops the other.
   */
  get numFmt(): string | undefined {
    return this.#numFmt;
  }

  set numFmt(numFmt: string | undefined) {
    this.#numFmt = numFmt;
  }

  /**
   * The cell's font — bold/italic/underline, size, colour, typeface — as a partial set
   * of the facets that differ from the default (only the facets actually set are carried,
   * exactly as OOXML stores them). `undefined` means the cell uses the workbook default
   * font. Like {@link fill} and {@link numFmt}, each cell owns its own font object, so a
   * font set on one cell never aliases or bleeds onto its row, column, or sheet siblings.
   */
  get font(): Partial<Font> | undefined {
    return this.#font;
  }

  set font(font: Partial<Font> | undefined) {
    this.#font = font;
  }
}
