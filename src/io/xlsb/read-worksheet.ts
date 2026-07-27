// The BIFF12 worksheet-body reader: one `xl/worksheets/sheetN.bin` in, one populated {@link Worksheet}
// out — the binary counterpart of `../xlsx/read-worksheet.ts`, filling the very same model.
//
// The cell table is a flat, row-major run of records: a `BrtRowHdr` opens a row and every cell record
// until the next one belongs to it, so the open row is a single variable rather than nested state.
// Each cell record is a `Cell` header followed by a value shaped by the record's own type — which is
// what makes the binary form quick to parse: there is no `t=` attribute to interpret, the record
// number *is* the type.
//
// **A formula cell surfaces its cached result, not its formula.** A BIFF12 formula is a `Ptg` token
// stream rather than text, and decoding it is a self-contained sub-project; until then a `BrtFmla*`
// record reads as the value Excel last computed — the same value the XML reader takes from `<v>`,
// without the `<f>` text beside it. A stated gap, not a silent one; see
// `docs/knowledge/specs/xlsb-binary-format-output.md`.

import {encodeAddress, MAX_COLUMN} from '../../core/address.ts';
import {isDateFormat, serialToDate} from '../../core/date.ts';
import {assignStyleFacets} from '../../core/style.ts';
import type {CellValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {applyXfToCell, type XfStyle} from '../xlsx/read-styles.ts';
import {errorCodeFor, RecordReader} from './primitives.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

// Every record that carries a cell. Membership drives the dispatch below, so a record type absent
// from this set is skipped whole rather than being mistaken for a cell and consuming the reader.
const CELL_RECORDS: ReadonlySet<number> = new Set([
  BRT.CellBlank,
  BRT.CellRk,
  BRT.CellError,
  BRT.CellBool,
  BRT.CellReal,
  BRT.CellSt,
  BRT.CellIsst,
  BRT.CellRString,
  BRT.FmlaString,
  BRT.FmlaNum,
  BRT.FmlaBool,
  BRT.FmlaError,
]);

/**
 * Read a worksheet part into `sheet`: its column and row geometry, its merged ranges, and every
 * non-empty cell with the style its index resolves to in `xfStyles`.
 */
export function parseWorksheet(
  part: Uint8Array,
  sheet: Worksheet,
  sharedStrings: readonly string[],
  xfStyles: ReadonlyArray<XfStyle>,
): void {
  // The open row, one-based as the model counts them. -1 means none is open, which a cell record
  // arriving before any row header (a malformed sheet) is dropped against rather than guessed at.
  let row = -1;
  // A row that declares a format supplies the default for its cells that carry none, as a column
  // does; the next row header replaces it.
  let rowStyle = -1;
  // A column's format is the last fallback. Column records always precede the cell table.
  const columnStyle = new Map<number, number>();
  // The sheet's default row height, in twips. Every row header restates its height whether or not the
  // row has one of its own, so the default is what tells the two apart — see {@link applyRow}.
  // `BrtWsFmtInfo` precedes the cell table, so it is always known by the time a row is read.
  let defaultRowHeight = -1;

  for (const record of readRecords(part)) {
    const reader = new RecordReader(record.data);
    if (record.type === BRT.WsFmtInfo) {
      reader.skip(6); // dxGCol, cchDefColWidth: the default *column* width, which the model does not read.
      defaultRowHeight = reader.u16();
    } else if (record.type === BRT.ColInfo) {
      applyColumn(reader, sheet, xfStyles, columnStyle);
    } else if (record.type === BRT.RowHdr) {
      const header = applyRow(reader, sheet, defaultRowHeight);
      row = header.row;
      rowStyle = header.styleIndex;
    } else if (record.type === BRT.MergeCell) {
      const {rowFirst, rowLast, colFirst, colLast} = reader.range();
      if (inGrid(colFirst, rowFirst) && inGrid(colLast, rowLast)) {
        sheet.mergeCells(
          `${encodeAddress(colFirst + 1, rowFirst + 1)}:${encodeAddress(colLast + 1, rowLast + 1)}`,
        );
      }
    } else if (CELL_RECORDS.has(record.type) && row > 0) {
      const {column, styleIndex} = reader.cell();
      if (!inGrid(column, row - 1)) continue;
      // A cell's own format wins, then its row's, then its column's — the order Excel applies.
      // Index 0 is the default xf, which BIFF12 writes where XML simply omits `s`, so it means
      // "no format of my own" and lets the row/column default through.
      const resolved =
        styleIndex > 0 ? styleIndex : rowStyle >= 0 ? rowStyle : (columnStyle.get(column) ?? -1);
      const style = resolved >= 0 ? xfStyles[resolved] : xfStyles[0];
      const cell = sheet.getCell(encodeAddress(column + 1, row));
      applyXfToCell(cell, style);
      cell.value = decodeCell(record.type, reader, sharedStrings, style?.numFmt);
    }
  }
}

// Excel's grid bounds, zero-based as the binary format counts. [MS-XLSB] states them as MUST
// constraints, which is exactly why a reader has to check them: a damaged or hostile file states
// whatever it likes, and an address beyond the grid has nowhere to go. Everything positional funnels
// through here before it reaches the model, so an out-of-grid record is dropped rather than turned
// into an unrepresentable address (which the address encoder would reject) or, worse, a column loop
// four billion iterations long.
const MAX_ROW_INDEX = 1048575;
const MAX_COLUMN_INDEX = MAX_COLUMN - 1;

function inGrid(column: number, row: number): boolean {
  return column >= 0 && column <= MAX_COLUMN_INDEX && row >= 0 && row <= MAX_ROW_INDEX;
}

// Decode a cell record's payload — the reader is positioned just past the shared `Cell` header, so
// what remains is exactly the value this record type carries.
function decodeCell(
  type: number,
  reader: RecordReader,
  sharedStrings: readonly string[],
  numFmt: string | undefined,
): CellValue {
  switch (type) {
    case BRT.CellRk:
      return asNumberOrDate(reader.rk(), numFmt);
    case BRT.CellReal:
    case BRT.FmlaNum:
      // A formula's cached numeric result honours the cell's date format exactly as a bare number
      // does, so a date-valued formula reads back as a Date rather than a serial.
      return asNumberOrDate(reader.f64(), numFmt);
    case BRT.CellBool:
    case BRT.FmlaBool:
      return reader.u8() !== 0;
    case BRT.CellError:
    case BRT.FmlaError: {
      const code = reader.u8();
      // An unrecognised error byte keeps the cell non-empty without inventing an error the model
      // does not define; there is no text form to fall back to as there is in XML.
      const error = errorCodeFor(code);
      return error === undefined ? null : {error};
    }
    case BRT.CellSt:
    case BRT.FmlaString:
      return reader.wideString();
    case BRT.CellRString:
      // Rich runs are not modelled in this cut; the flattened text is what a consumer sees.
      return reader.richString();
    case BRT.CellIsst:
      return sharedStrings[reader.u32()] ?? '';
    default:
      // BrtCellBlank: formatted but empty. The style is already applied; the value is genuinely none.
      return null;
  }
}

// A number stored under a date format is a date serial — surface it as a Date so a date read from an
// `.xlsb` is the same value the `.xlsx` twin yields, not a bare number.
function asNumberOrDate(value: number, numFmt: string | undefined): CellValue {
  return numFmt !== undefined && isDateFormat(numFmt) ? serialToDate(value) : value;
}

// `BrtRowHdr` ([MS-XLSB] 2.4.770): the row index, its default format, its height, and a byte of
// layout flags. Returns the open row (one-based) and the style index its cells inherit.
function applyRow(
  reader: RecordReader,
  sheet: Worksheet,
  defaultRowHeight: number,
): {row: number; styleIndex: number} {
  const index = reader.u32();
  const styleIndex = reader.u32();
  const height = reader.u16();
  reader.skip(1); // fExtraAsc/fExtraDsc: border padding, a rendering hint the model does not carry.
  const flags = reader.u8();
  // A row beyond the grid closes the open row without opening another, so its cells are dropped too.
  if (index > MAX_ROW_INDEX) return {row: -1, styleIndex: -1};

  const row = index + 1;
  const properties = sheet.getRow(row);
  // Every row header restates a height; only a row whose height is its *own* has one to record. That
  // is a row the user sized by hand, or one Excel auto-fitted to a taller font or wrapped text — both
  // differ from the sheet default, which is exactly when XML emits `ht`. A row merely restating the
  // default carries no height, so it must not read back with one.
  if ((flags & ROW_CUSTOM_HEIGHT) !== 0 || height !== defaultRowHeight) {
    properties.height = height / TWIPS_PER_POINT;
  }
  if ((flags & ROW_HIDDEN) !== 0) properties.hidden = true;
  const outlineLevel = flags & ROW_OUTLINE_LEVEL;
  if (outlineLevel > 0) properties.outlineLevel = outlineLevel;
  if ((flags & ROW_COLLAPSED) !== 0) properties.collapsed = true;
  // The row's format applies only when it says so, mirroring XML's `customFormat="1"` gate.
  return {row, styleIndex: (flags & ROW_CUSTOM_FORMAT) !== 0 ? styleIndex : -1};
}

const TWIPS_PER_POINT = 20;
const ROW_OUTLINE_LEVEL = 0b0000_0111;
const ROW_COLLAPSED = 0b0000_1000;
const ROW_HIDDEN = 0b0001_0000;
const ROW_CUSTOM_HEIGHT = 0b0010_0000;
const ROW_CUSTOM_FORMAT = 0b0100_0000;

// `BrtColInfo` ([MS-XLSB] 2.4.319): one record per run of identically-sized columns.
function applyColumn(
  reader: RecordReader,
  sheet: Worksheet,
  xfStyles: ReadonlyArray<XfStyle>,
  columnStyle: Map<number, number>,
): void {
  const first = reader.u32();
  const last = reader.u32();
  const width = reader.u32();
  const styleIndex = reader.u32();
  const flags = reader.u16();
  const style = styleIndex > 0 ? xfStyles[styleIndex] : undefined;
  const outlineLevel = (flags >> COLUMN_OUTLINE_SHIFT) & 0b111;

  // A run can name every column to the right of the data; materialising all 16 384 of them would
  // turn a two-column sheet into a 16 384-entry model, so a run is only applied where it carries
  // something a default column does not.
  if (
    style === undefined &&
    outlineLevel === 0 &&
    (flags & (COLUMN_HIDDEN | COLUMN_CUSTOM_WIDTH | COLUMN_COLLAPSED)) === 0
  ) {
    return;
  }
  // The loop bound comes from the file, so it is clamped to the grid before it is one: an unclamped
  // run declaring four billion columns is a denial of service, not a wide sheet.
  const lastInGrid = Math.min(last, MAX_COLUMN_INDEX);
  if (first > lastInGrid) return;
  for (let index = first; index <= lastInGrid; index++) {
    const properties = sheet.getColumn(index + 1);
    // The stored width is taken whether or not the file marks it user-set, matching the XML reader:
    // a `<col>`/`BrtColInfo` exists only for a column that differs from the sheet default in *some*
    // way, and it always states the width that column actually has.
    properties.width = width / COLUMN_WIDTH_UNITS;
    if ((flags & COLUMN_HIDDEN) !== 0) properties.hidden = true;
    if (outlineLevel > 0) properties.outlineLevel = outlineLevel;
    if ((flags & COLUMN_COLLAPSED) !== 0) properties.collapsed = true;
    if (style !== undefined) assignStyleFacets(properties, style);
    if (styleIndex > 0) columnStyle.set(index, styleIndex);
  }
}

// Column width is stored in 1/256ths of a character, where XML states the character count directly.
const COLUMN_WIDTH_UNITS = 256;
const COLUMN_HIDDEN = 0b0000_0001;
const COLUMN_CUSTOM_WIDTH = 0b0000_0010;
const COLUMN_COLLAPSED = 0b1000_0000_0000;
const COLUMN_OUTLINE_SHIFT = 8;
