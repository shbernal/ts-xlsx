// The sheet's used-range extent: how far the grid reaches, maintained as the sheet is built rather
// than re-derived on every read.
//
// The naive reading is a scan of every cell, and it is quadratic exactly where it hurts most. The
// three appenders (addRows, addColumns, and the streaming writer's row numbering) each read the
// extent once per line they append, so appending row by row visits O(rows^2) cells: 16k rows took
// 15 seconds where the same rows through one bulk call took 85 ms.
//
// Remembering the answer is not open to us. A caller holding a `Cell` can style, note or clear it
// without the sheet ever hearing about it, so a cached extent can go stale invisibly, and a wrong
// used range is worse than a slow one: it is what lets an append land on a row someone prepared.
//
// What *is* exactly maintainable is the structure. A cell only enters the grid through the sheet,
// so the highest row key, the highest materialised column index, and the highest declared line are
// facts the sheet observes as they happen, and every one of them bounds the extent from above. In
// the case this exists for, the top line was just appended and is plainly used, so the read becomes
// "is the topmost line used?" rather than "which line is topmost?". When it is not (a cell
// materialised by `getCell` and never filled), the answer falls back to the full scan, which is
// where it started. The bounds may therefore overstate; they may never understate, which is why
// every edit that can shrink the grid marks them stale instead of adjusting them.

import {type Cell, cellCarriesContent} from './cell.ts';
import type {MergeRect} from './merge.ts';

/**
 * The sheet storage an extent reads. Held by reference and never reassigned, so the extent always
 * sees the live grid; the value types are irrelevant to it, only which keys are present.
 */
export interface ExtentStorage {
  readonly rows: ReadonlyMap<number, ReadonlyMap<number, Cell>>;
  readonly rowProperties: ReadonlyMap<number, unknown>;
  readonly columns: ReadonlyMap<number, unknown>;
  readonly mergeRects: readonly MergeRect[];
}

/** Whether any cell materialised in a row is used, short-circuiting on the first one that is. */
function rowIsUsed(cols: ReadonlyMap<number, Cell>): boolean {
  for (const cell of cols.values()) {
    if (cellCarriesContent(cell)) return true;
  }
  return false;
}

/**
 * The last used row and column of a sheet, tracked incrementally. The sheet reports every edit that
 * can move the grid's outer keys; this turns those reports into the two accessors behind
 * `Worksheet.rowCount` and `Worksheet.columnCount`.
 */
export class UsedExtent {
  readonly #storage: ExtentStorage;

  // The highest row key the grid holds, and the highest column index materialised in any row. Upper
  // bounds on the extent, never the extent itself: the line they name may hold nothing but blank
  // cells, so each is confirmed against the grid before it is believed.
  #topRowKey = 0;
  #topColumnKey = 0;

  // The highest line that bounds the used range by declaration alone: a row height or outline level,
  // a column width, a merge reaching past the last value. Unlike the keys above these are taken
  // without asking whether anything in the line is used, so they must be exact rather than merely
  // bounding.
  #topDeclaredRow = 0;
  #topDeclaredColumn = 0;

  // Whether an edit that can shrink the grid has landed since the four above were last known good.
  #stale = false;

  constructor(storage: ExtentStorage) {
    this.#storage = storage;
  }

  /** Note a cell materialised at `(row, col)`. Whether it goes on to carry anything is not this
   * class's business: the bound is on where a used cell *could* be, and the read confirms it. */
  noteCell(row: number, col: number): void {
    if (row > this.#topRowKey) this.#topRowKey = row;
    if (col > this.#topColumnKey) this.#topColumnKey = col;
  }

  /** Note a row that bounds the used range by declaration: one given properties of its own. */
  noteDeclaredRow(row: number): void {
    if (row > this.#topDeclaredRow) this.#topDeclaredRow = row;
  }

  /** Note a column that bounds the used range by declaration: one given properties of its own. */
  noteDeclaredColumn(col: number): void {
    if (col > this.#topDeclaredColumn) this.#topDeclaredColumn = col;
  }

  /** Note a merged region, which occupies its whole rectangle and so bounds both axes. */
  noteMerge(rect: MergeRect): void {
    this.noteDeclaredRow(rect.bottom);
    this.noteDeclaredColumn(rect.right);
  }

  /**
   * Report an edit that can pull the grid's outer keys inward: an eviction, a splice, an unmerge, a
   * row replaced wholesale. The bounds are recomputed once, on the next read, rather than adjusted
   * here, because finding the *new* maximum after a deletion is the same scan as finding it from
   * scratch. Appending never invalidates, which is the whole point.
   */
  invalidate(): void {
    this.#stale = true;
  }

  /** Report that the grid has been emptied, which puts the bounds at a known zero rather than
   * stale: nothing is left to scan. */
  reset(): void {
    this.#topRowKey = 0;
    this.#topColumnKey = 0;
    this.#topDeclaredRow = 0;
    this.#topDeclaredColumn = 0;
    this.#stale = false;
  }

  /** The 1-based index of the last row carrying anything, or 0 when the sheet holds nothing. */
  get lastRow(): number {
    this.#refresh();
    if (this.#topRowKey <= this.#topDeclaredRow) return this.#topDeclaredRow;
    const top = this.#storage.rows.get(this.#topRowKey);
    if (top !== undefined && rowIsUsed(top)) return this.#topRowKey;
    let last = this.#topDeclaredRow;
    for (const [number, cols] of this.#storage.rows) {
      if (number > last && rowIsUsed(cols)) last = number;
    }
    return last;
  }

  /** The 1-based index of the last column carrying anything, or 0 when the sheet holds nothing. */
  get lastColumn(): number {
    this.#refresh();
    if (this.#topColumnKey <= this.#topDeclaredColumn) return this.#topDeclaredColumn;
    // One lookup per row rather than a walk of every cell: only the topmost column is in question,
    // and a Map.get against each row answers it.
    for (const cols of this.#storage.rows.values()) {
      const cell = cols.get(this.#topColumnKey);
      if (cell !== undefined && cellCarriesContent(cell)) return this.#topColumnKey;
    }
    let last = this.#topDeclaredColumn;
    for (const cols of this.#storage.rows.values()) {
      for (const [col, cell] of cols) {
        if (col > last && cellCarriesContent(cell)) last = col;
      }
    }
    return last;
  }

  // Re-derive the four bounds from the live grid after a shrinking edit. O(cells), paid once per
  // such edit rather than once per read.
  #refresh(): void {
    if (!this.#stale) return;
    this.#stale = false;
    this.#topRowKey = 0;
    this.#topColumnKey = 0;
    this.#topDeclaredRow = 0;
    this.#topDeclaredColumn = 0;
    for (const [number, cols] of this.#storage.rows) {
      if (number > this.#topRowKey) this.#topRowKey = number;
      for (const col of cols.keys()) {
        if (col > this.#topColumnKey) this.#topColumnKey = col;
      }
    }
    for (const number of this.#storage.rowProperties.keys()) this.noteDeclaredRow(number);
    for (const index of this.#storage.columns.keys()) this.noteDeclaredColumn(index);
    for (const rect of this.#storage.mergeRects) this.noteMerge(rect);
  }
}
