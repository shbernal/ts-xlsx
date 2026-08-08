// How a `RowInput` is read: the one interpretation of "a row's worth of values" that every authoring
// entry point on `Worksheet` shares — `addRow`, `addRows`, `insertRow`, `spliceRows`.
//
// This is not grid mechanics, which is why it is not in `grid-edits.ts`: nothing here shifts or
// re-anchors anything, and the arithmetic is indifferent to where the row ends up. It is the public
// API's *input vocabulary* — positional array versus key-addressed object, and what a hole in either
// one means — held in one place so appending into the live grid and splicing into a detached row can
// never drift on the answer. A row that placed values differently depending on which method received
// it would be the kind of bug no single call site looks wrong for.

import {AuthoringError} from '../errors.ts';
import {Cell} from './cell.ts';
import type {CellValue} from './value.ts';
import type {ColumnProperties, RowInput} from './worksheet.ts';

/**
 * Resolve a `RowInput` to the (1-based column, value) placements it names. A positional array maps
 * each value to its column from A, skipping a hole or an explicit `undefined` so that column is left
 * untouched; a keyed object maps each value under the column carrying the matching key.
 *
 * `Array.isArray`, not `instanceof Array`: a row built in another realm (a vm context, a browser
 * iframe) is still an array but fails the identity check, and would then be walked as a keyed object —
 * placing nothing.
 */
export function rowPlacements(
  values: RowInput,
  columns: ReadonlyMap<number, ColumnProperties>,
): Array<[number, CellValue]> {
  if (Array.isArray(values)) {
    const placements: Array<[number, CellValue]> = [];
    values.forEach((value, index) => {
      if (value !== undefined) placements.push([index + 1, value]);
    });
    return placements;
  }
  return Object.entries(values).map(([key, value]) => [columnIndexByKey(columns, key), value]);
}

/**
 * Build the detached cell row an insert introduces: a fresh cell per placement, positioned at `number`,
 * keyed by column. The grid-edit machinery then splices this map into place.
 */
export function buildRowCells(
  number: number,
  values: RowInput,
  columns: ReadonlyMap<number, ColumnProperties>,
): Map<number, Cell> {
  const row = new Map<number, Cell>();
  for (const [col, value] of rowPlacements(values, columns)) {
    const cell = new Cell(number, col);
    cell.value = value;
    row.set(col, cell);
  }
  return row;
}

/**
 * The 1-based index of the column carrying `key` (see {@link ColumnProperties.key}).
 *
 * @throws {AuthoringError} if no column declares that key.
 */
export function columnIndexByKey(
  columns: ReadonlyMap<number, ColumnProperties>,
  key: string,
): number {
  for (const [index, properties] of columns) {
    if (properties.key === key) return index;
  }
  throw new AuthoringError(
    `no column is keyed ${JSON.stringify(key)} — set getColumn(n).key first`,
  );
}
