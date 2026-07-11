// A single cell: a typed value at a fixed 1-based `(row, col)` position.
//
// Position is immutable and numeric — the corpus locks `col`/`row` as 1-based
// *numbers* (legacy shipped a type declaration calling them strings, which broke
// strict consumers). The value is the only mutable state here; assigning it routes
// through the value model so the cell's `type` is always consistent with what it holds.

import {encodeAddress} from './address.ts';
import {type CellValue, type ValueType, coerceCellValue, detectValueType} from './value.ts';

export class Cell {
  /** 1-based row index. */
  readonly row: number;
  /** 1-based column index. */
  readonly col: number;

  #value: CellValue = null;

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
}
