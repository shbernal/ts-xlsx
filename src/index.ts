// Public entry point for the ts-xlsx rewrite (Phase 3).
//
// This barrel is intentionally thin and provisional: the definitive public API is
// designed module by module as the rewrite lands, corpus-first. Only surfaces that
// are implemented, strict-typed, and corpus-backed are re-exported here.

export {
  type CellAddress,
  columnToNumber,
  decodeAddress,
  decodeRange,
  encodeAddress,
  MAX_COLUMN,
  numberToColumn,
  type RangeAddress,
} from './core/address.ts';
export {Cell} from './core/cell.ts';
export {
  type AnchoredImage,
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  isOneCellAnchor,
  type OneCellAnchor,
  PX_TO_EMU,
  type TwoCellAnchor,
  type WorkbookImage,
} from './core/image.ts';
export type {HeaderFooter, PageMargins} from './core/page-setup.ts';
export {
  type PivotMetric,
  PivotTable,
  type PivotTableOptions,
} from './core/pivot-table.ts';
export type {
  SheetProtection,
  SheetProtectionCredential,
  SheetProtectionFlags,
  SheetProtectionOptions,
} from './core/protection.ts';
export type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  Color,
  Fill,
  FillPatternType,
  Font,
  FontVerticalAlignment,
  GradientFill,
  GradientStop,
  HorizontalAlignment,
  PatternFill,
  Protection,
  UnderlineStyle,
  VerticalAlignment,
} from './core/style.ts';
export {
  Table,
  type TableColumn,
  type TableOptions,
  type TableRegion,
} from './core/table.ts';
export {
  type CellValue,
  coerceCellValue,
  detectValueType,
  ERROR_CODES,
  type ErrorCode,
  type ErrorValue,
  type FormulaResult,
  type FormulaValue,
  type HyperlinkValue,
  isErrorCode,
  type RichTextRun,
  type RichTextValue,
  richTextToPlain,
  type SharedFormulaValue,
  ValueType,
} from './core/value.ts';
export {
  type AddWorksheetOptions,
  Workbook,
  type WorkbookProperties,
} from './core/workbook.ts';
export type {WorkbookProtection} from './core/workbook-protection.ts';
export {
  type ColumnProperties,
  type RowProperties,
  Worksheet,
  type WorksheetProperties,
  type WorksheetState,
} from './core/worksheet.ts';
export {type CsvReadOptions, readCsv} from './io/csv/read.ts';
export {type CsvWriteOptions, writeCsv, writeCsvText} from './io/csv/write.ts';
export {type ReadXlsxOptions, readXlsx} from './io/xlsx/read.ts';
export {
  type ReadSheetRowsOptions,
  readSheetRows,
  readWorkbookStream,
  type StreamedCell,
  type StreamedCellStyle,
  type StreamedRow,
  type StreamedSheet,
} from './io/xlsx/read-rows.ts';
export {writeXlsx} from './io/xlsx/write.ts';
