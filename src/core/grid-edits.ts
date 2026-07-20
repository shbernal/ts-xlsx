// Structural-edit machinery: the splice arithmetic that inserts or deletes whole rows and columns
// and keeps everything anchored to the grid moving in step — line metadata, merged ranges, tables,
// anchored images, and shared-formula clones. It is isolated from Worksheet because it is pure grid
// mechanics: it holds the sheet's storage containers by reference and mutates them in place, and
// touches none of the public cell API. Worksheet builds the cells an insert introduces, then hands
// the pre-built rows (or the raw column values) here for the shift.

import {decodeAddress, decodeRange, encodeAddress} from './address.ts';
import {Cell, copyCellContent} from './cell.ts';
import {type AnchoredImage, type AnchorPoint, type ImageAnchor, isOneCellAnchor} from './image.ts';
import type {MergeRect} from './merge.ts';
import type {Table} from './table.ts';
import {type CellValue, isSharedFormulaValue, type SharedFormulaValue} from './value.ts';
import type {ColumnProperties, RowProperties} from './worksheet.ts';

// The shift rule shared by every re-anchoring pass: a coordinate before the edit stays put, one at or
// after the edited span shifts by `delta`, and one inside a deleted span clamps to the cut line (`start`).
function shiftIndex(v: number, start: number, count: number, delta: number): number {
  return v < start ? v : v >= start + count ? v + delta : start;
}

// The sheet's mutable storage, shared by reference with Worksheet. Never reassigned — only mutated in
// place — so the two views stay in sync through every splice.
interface GridStorage {
  readonly rows: Map<number, Map<number, Cell>>;
  readonly rowProperties: Map<number, RowProperties>;
  readonly columns: Map<number, ColumnProperties>;
  readonly merges: string[];
  readonly mergeRects: MergeRect[];
  readonly tables: Table[];
  readonly images: AnchoredImage[];
}

export class GridEdits {
  readonly #rows: Map<number, Map<number, Cell>>;
  readonly #rowProperties: Map<number, RowProperties>;
  readonly #columns: Map<number, ColumnProperties>;
  readonly #merges: string[];
  readonly #mergeRects: MergeRect[];
  readonly #tables: Table[];
  readonly #images: AnchoredImage[];

  constructor(storage: GridStorage) {
    this.#rows = storage.rows;
    this.#rowProperties = storage.rowProperties;
    this.#columns = storage.columns;
    this.#merges = storage.merges;
    this.#mergeRects = storage.mergeRects;
    this.#tables = storage.tables;
    this.#images = storage.images;
  }

  // Apply a delete-then-insert to the row grid: surviving rows below the edit shift by
  // `inserted.length - count`, deleted rows drop out, and the pre-built inserted rows land at `start`.
  // Row metadata and merged ranges shift the same way, so a formatting-only row or a covered merge
  // stays aligned with the data it describes.
  spliceRows(start: number, count: number, inserted: Map<number, Cell>[]): void {
    const delta = inserted.length - count;
    const shifted = new Map<number, Map<number, Cell>>();
    for (const [row, cols] of this.#rows) {
      if (row < start) shifted.set(row, cols);
      else if (row >= start + count) shifted.set(row + delta, this.#relocateRow(cols, row + delta));
    }
    inserted.forEach((cols, i) => {
      shifted.set(start + i, this.#relocateRow(cols, start + i));
    });
    this.#rows.clear();
    for (const [row, cols] of shifted) this.#rows.set(row, cols);

    this.#shiftLineProperties(this.#rowProperties, start, count, delta);
    this.#shiftMerges('row', start, count, delta);
    this.#shiftTables('row', start, count, delta);
    this.#shiftImages('row', start, count, delta);
    this.#reanchorSharedFormulas('row', start, count, delta);
  }

  // Apply a delete-then-insert to the column grid: cells left of the edit stay, cells at or beyond the
  // deleted span shift by `inserts.length - count` carrying their content, and the inserted column
  // values materialise as fresh cells at `start`. Column metadata, merges, tables, images, and
  // shared-formula clones re-anchor the same way.
  spliceColumns(start: number, count: number, inserts: CellValue[][]): void {
    const delta = inserts.length - count;
    for (const [row, cols] of this.#rows) {
      const shifted = new Map<number, Cell>();
      for (const [col, cell] of cols) {
        if (col < start) {
          shifted.set(col, cell);
        } else if (col >= start + count) {
          const dest = col + delta;
          const moved = new Cell(row, dest);
          copyCellContent(cell, moved);
          shifted.set(dest, moved);
        }
      }
      inserts.forEach((values, i) => {
        const value = values[row - 1];
        if (value !== undefined) {
          const cell = new Cell(row, start + i);
          cell.value = value;
          shifted.set(start + i, cell);
        }
      });
      this.#rows.set(row, shifted);
    }
    this.#shiftLineProperties(this.#columns, start, count, delta);
    this.#shiftMerges('col', start, count, delta);
    this.#shiftTables('col', start, count, delta);
    this.#shiftImages('col', start, count, delta);
    this.#reanchorSharedFormulas('col', start, count, delta);
  }

  // Rebuild a row's cells at a new row index. `Cell` fixes its position at construction, so a moved
  // row is a fresh set of cells at `destRow` carrying the originals' content.
  #relocateRow(cols: Map<number, Cell>, destRow: number): Map<number, Cell> {
    const moved = new Map<number, Cell>();
    for (const [col, cell] of cols) {
      if (cell.row === destRow) {
        moved.set(col, cell);
      } else {
        const copy = new Cell(destRow, col);
        copyCellContent(cell, copy);
        moved.set(col, copy);
      }
    }
    return moved;
  }

  // Re-anchor shared-formula clones through a splice on the given axis. A clone stores its master's
  // absolute address; when the splice shifts the master, that stored address goes stale and the writer
  // would reject the clone as orphaned. Applying the same shift the grid used keeps each clone pointed
  // at its master's new cell. A master whose axis coordinate falls in the deleted span clamps to the
  // cut line like a merge edge — a genuinely orphaned clone the writer then reports legibly.
  #reanchorSharedFormulas(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        if (!isSharedFormulaValue(value)) continue;
        const master = decodeAddress(value.sharedFormula);
        if (master.col === undefined || master.row === undefined) continue;
        const anchored =
          axis === 'row'
            ? encodeAddress(master.col, shiftIndex(master.row, start, count, delta))
            : encodeAddress(shiftIndex(master.col, start, count, delta), master.row);
        if (anchored === value.sharedFormula) continue;
        const reanchored: SharedFormulaValue = {...value, sharedFormula: anchored};
        cell.value = reanchored;
      }
    }
  }

  // Shift a line-metadata map (row properties keyed by row, or column properties keyed by column)
  // through a splice: entries before the edit stay, entries within the deleted span drop, entries
  // after shift by `delta`. Mutates the map in place.
  #shiftLineProperties<T>(map: Map<number, T>, start: number, count: number, delta: number): void {
    const shifted = new Map<number, T>();
    for (const [index, value] of map) {
      if (index < start) shifted.set(index, value);
      else if (index >= start + count) shifted.set(index + delta, value);
    }
    map.clear();
    for (const [index, value] of shifted) map.set(index, value);
  }

  // Re-anchor merged ranges through a row or column splice. A range wholly before the edit is
  // untouched; one wholly after shifts by `nInserts - count`; one whose covered rows/columns are
  // entirely deleted is dropped. A range straddling the cut is a genuinely ambiguous geometry — its
  // edges are clamped to the cut line as a best effort. Unbounded whole-row/column merges carry no
  // rectangle and pass through unchanged.
  #shiftMerges(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const shift = (v: number): number => shiftIndex(v, start, count, delta);
    const merges: string[] = [];
    const rects: MergeRect[] = [];
    for (const range of this.#merges) {
      const {top, left, bottom, right} = decodeRange(range);
      if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
        merges.push(range);
        continue;
      }
      const [lo, hi] = axis === 'row' ? [top, bottom] : [left, right];
      if (lo >= start && hi < start + count) continue;
      const rect: MergeRect =
        axis === 'row'
          ? {top: shift(top), left, bottom: shift(bottom), right}
          : {top, left: shift(left), bottom, right: shift(right)};
      rects.push(rect);
      merges.push(
        `${encodeAddress(rect.left, rect.top)}:${encodeAddress(rect.right, rect.bottom)}`,
      );
    }
    this.#merges.length = 0;
    this.#merges.push(...merges);
    this.#mergeRects.length = 0;
    this.#mergeRects.push(...rects);
  }

  // Re-pin the sheet's tables through a splice on the given axis, dropping any table a delete leaves
  // with no row to occupy. `Table` owns the shift arithmetic; the sheet only prunes the casualties.
  #shiftTables(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const survivors = this.#tables.filter((table) =>
      axis === 'row'
        ? table.shiftRows(start, count, delta)
        : table.shiftColumns(start, count, delta),
    );
    this.#tables.length = 0;
    this.#tables.push(...survivors);
  }

  // Re-pin anchored images through a splice. An anchor point moves like a merge edge: a point before
  // the cut stays, one at or after it shifts by `delta`, and one inside a deleted span clamps to the
  // cut line. Grid points are 0-based, so each is converted to the 1-based coordinate the shared
  // shift arithmetic uses and back. An anchor whose points both move keeps its size; an anchor
  // straddling the cut grows or shrinks, matching how Excel reflows a picture across inserted rows.
  #shiftImages(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const shiftPoint = (point: AnchorPoint): AnchorPoint => {
      const zeroBased = axis === 'row' ? point.row : point.col;
      const shifted = shiftIndex(zeroBased + 1, start, count, delta) - 1;
      if (shifted === zeroBased) return point;
      return axis === 'row' ? {...point, row: shifted} : {...point, col: shifted};
    };
    const moved: AnchoredImage[] = this.#images.map((image) => {
      const from = shiftPoint(image.anchor.from);
      const anchor: ImageAnchor = isOneCellAnchor(image.anchor)
        ? {...image.anchor, from}
        : {...image.anchor, from, to: shiftPoint(image.anchor.to)};
      return {imageId: image.imageId, anchor};
    });
    this.#images.length = 0;
    this.#images.push(...moved);
  }
}
