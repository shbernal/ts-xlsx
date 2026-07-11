// A worksheet: a sparsely-populated grid of cells addressed by A1 reference.
//
// Storage is sparse by construction — a spreadsheet is mostly empty, so cells
// materialise on first access and only occupied positions cost memory. Rows,
// columns, merges, and views layer on in later slices; this slice is the cell grid
// and the addressing that reaches into it.

import {decodeAddress} from './address.ts';
import {Cell} from './cell.ts';

export interface WorksheetState {
  /** Sheet visibility, as Excel models it. Defaults to `visible`. */
  readonly state: 'visible' | 'hidden' | 'veryHidden';
}

export class Worksheet {
  readonly name: string;
  /** 1-based workbook-assigned id, stable for the sheet's lifetime. */
  readonly id: number;
  state: WorksheetState['state'];

  // Row-major sparse storage: row index → (column index → cell). Keeping rows as the
  // outer key makes whole-row iteration cheap once rows land, and mirrors how OOXML
  // serializes (`<row>` wrapping `<c>`).
  readonly #rows = new Map<number, Map<number, Cell>>();

  constructor(name: string, id: number, state: WorksheetState['state'] = 'visible') {
    this.name = name;
    this.id = id;
    this.state = state;
  }

  /**
   * Get the cell at an A1 reference, creating it on first access. The reference must
   * name both a column and a row (`"B3"`); a whole-row or whole-column reference is
   * not a cell and is rejected.
   *
   * @throws {SyntaxError} if the reference does not resolve to a single cell.
   */
  getCell(reference: string): Cell {
    const {col, row} = decodeAddress(reference);
    if (col === undefined || row === undefined) {
      throw new SyntaxError(`"${reference}" is not a single-cell reference — it omits a column or row`);
    }
    return this.#cellAt(row, col);
  }

  /** Whether a cell has been materialised at the given 1-based position. */
  hasCell(row: number, col: number): boolean {
    return this.#rows.get(row)?.has(col) ?? false;
  }

  #cellAt(row: number, col: number): Cell {
    let cols = this.#rows.get(row);
    if (cols === undefined) {
      cols = new Map<number, Cell>();
      this.#rows.set(row, cols);
    }
    let cell = cols.get(col);
    if (cell === undefined) {
      cell = new Cell(row, col);
      cols.set(col, cell);
    }
    return cell;
  }
}
