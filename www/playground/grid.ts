/**
 * A worksheet as rows of display cells.
 *
 * The bound lives here rather than in the component, because a bound applied while painting
 * is a bound that has already read four thousand rows. The model decides how much of a sheet
 * a reader is shown, and reports what it left out so the page can say so.
 *
 * This is the *value's* text, not Excel's rendering of it: the number format is carried
 * through as data for the panel to display, and never applied. The library models formats,
 * it does not implement a formatting engine, and a grid that quietly invented one would
 * be showing a workbook this library cannot actually produce.
 *
 * Nothing here touches a DOM.
 */

import {
  type Alignment,
  type CellStyle,
  cellValueToText,
  decodeRange,
  type Fill,
  type Font,
  numberToColumn,
  type ValueType,
  detectValueType,
  type Worksheet,
} from '../../src/index.ts';

export interface DisplayCell {
  readonly row: number;
  readonly col: number;
  readonly address: string;
  /** The value's text. Empty for an empty cell. */
  readonly text: string;
  readonly type: ValueType;
  readonly numFmt: string | undefined;
  readonly font: Font | undefined;
  readonly fill: Fill | undefined;
  readonly alignment: Alignment | undefined;
  /** Set on the top-left cell of a merge, and the reason the cells it covers are absent. */
  readonly merge: {readonly rowSpan: number; readonly colSpan: number} | undefined;
}

export interface DisplayRow {
  readonly number: number;
  readonly cells: readonly DisplayCell[];
}

export interface Grid {
  readonly sheetName: string;
  /** Column headers, `A`, `B`, `C`, for the columns actually shown. */
  readonly columns: readonly string[];
  readonly rows: readonly DisplayRow[];
  readonly totalRows: number;
  readonly totalColumns: number;
  /** How many rows and columns exist beyond the window, so the page can say so honestly. */
  readonly hiddenRows: number;
  readonly hiddenColumns: number;
}

export interface GridBounds {
  readonly maxRows: number;
  readonly maxColumns: number;
}

export const DEFAULT_BOUNDS: GridBounds = {maxRows: 200, maxColumns: 26};

export function toGrid(sheet: Worksheet, bounds: GridBounds = DEFAULT_BOUNDS): Grid {
  const used = sheet.usedRange;
  if (used === undefined) {
    return {
      sheetName: sheet.name,
      columns: [],
      rows: [],
      totalRows: 0,
      totalColumns: 0,
      hiddenRows: 0,
      hiddenColumns: 0,
    };
  }

  const totalRows = used.bottom - used.top + 1;
  const totalColumns = used.right - used.left + 1;
  const shownRows = Math.min(totalRows, Math.max(0, bounds.maxRows));
  const shownColumns = Math.min(totalColumns, Math.max(0, bounds.maxColumns));

  const covered = mergeCover(sheet);
  const rows: DisplayRow[] = [];
  for (let offset = 0; offset < shownRows; offset++) {
    const row = used.top + offset;
    const cells: DisplayCell[] = [];
    for (let span = 0; span < shownColumns; span++) {
      const col = used.left + span;
      // A cell swallowed by a merge is not rendered at all; the anchor's span covers it.
      if (covered.hidden.has(key(row, col))) continue;
      cells.push(displayCell(sheet, row, col, covered.spans.get(key(row, col))));
    }
    rows.push({number: row, cells});
  }

  const columns: string[] = [];
  for (let span = 0; span < shownColumns; span++) columns.push(numberToColumn(used.left + span));

  return {
    sheetName: sheet.name,
    columns,
    rows,
    totalRows,
    totalColumns,
    hiddenRows: totalRows - shownRows,
    hiddenColumns: totalColumns - shownColumns,
  };
}

function displayCell(
  sheet: Worksheet,
  row: number,
  col: number,
  merge: {readonly rowSpan: number; readonly colSpan: number} | undefined,
): DisplayCell {
  const address = `${numberToColumn(col)}${row}`;
  // `hasCell` first: `getCell` materialises a cell, and walking a window of a large sheet
  // would otherwise leave the workbook bigger than the read that produced it.
  if (!sheet.hasCell(row, col)) {
    return {
      row,
      col,
      address,
      text: '',
      type: 'null',
      numFmt: undefined,
      font: undefined,
      fill: undefined,
      alignment: undefined,
      merge,
    };
  }
  const cell = sheet.getCell(address);
  const style: CellStyle = cell.style;
  return {
    row,
    col,
    address,
    text: cellValueToText(cell.value),
    type: detectValueType(cell.value),
    numFmt: style.numFmt,
    font: style.font,
    fill: style.fill,
    alignment: style.alignment,
    merge,
  };
}

/** A cell's identity in the two maps below. */
function key(row: number, col: number): string {
  return `${row}:${col}`;
}

interface MergeCover {
  /** Anchor cell to its span. */
  readonly spans: Map<string, {readonly rowSpan: number; readonly colSpan: number}>;
  /** Every cell a merge covers except its anchor. */
  readonly hidden: Set<string>;
}

function mergeCover(sheet: Worksheet): MergeCover {
  const spans = new Map<string, {rowSpan: number; colSpan: number}>();
  const hidden = new Set<string>();
  for (const reference of sheet.merges) {
    const rect = boundedRect(reference);
    if (rect === undefined) continue;
    spans.set(key(rect.top, rect.left), {
      rowSpan: rect.bottom - rect.top + 1,
      colSpan: rect.right - rect.left + 1,
    });
    for (let row = rect.top; row <= rect.bottom; row++) {
      for (let col = rect.left; col <= rect.right; col++) {
        if (row !== rect.top || col !== rect.left) hidden.add(key(row, col));
      }
    }
  }
  return {spans, hidden};
}

interface Rect {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/**
 * A range with all four edges, or nothing.
 *
 * A range address leaves an edge unset for a whole-row or whole-column reference, which a
 * merge cannot be. Returning nothing is the honest answer for one that arrives anyway:
 * inventing the missing bound would blank out cells nothing merged.
 */
function boundedRect(reference: string): Rect | undefined {
  const {top, left, bottom, right} = decodeRange(reference);
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    return undefined;
  }
  return {top, left, bottom, right};
}
