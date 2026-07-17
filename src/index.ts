// Public entry point for the ts-xlsx rewrite (Phase 3).
//
// This barrel is intentionally thin and provisional: the definitive public API is
// designed module by module as the rewrite lands, corpus-first. Only surfaces that
// are implemented, strict-typed, and corpus-backed are re-exported here.

export {
  type CellAddress,
  type RangeAddress,
  MAX_COLUMN,
  columnToNumber,
  numberToColumn,
  decodeAddress,
  decodeRange,
  encodeAddress,
} from './core/address.ts';

export {
  type CellValue,
  type ErrorCode,
  type ErrorValue,
  type FormulaResult,
  type FormulaValue,
  type SharedFormulaValue,
  type RichTextRun,
  type RichTextValue,
  type HyperlinkValue,
  ERROR_CODES,
  ValueType,
  coerceCellValue,
  detectValueType,
  isErrorCode,
} from './core/value.ts';

export {
  type Alignment,
  type Border,
  type BorderEdge,
  type BorderStyle,
  type Color,
  type Fill,
  type FillPatternType,
  type Font,
  type FontVerticalAlignment,
  type HorizontalAlignment,
  type PatternFill,
  type Protection,
  type UnderlineStyle,
  type VerticalAlignment,
} from './core/style.ts';

export {
  type AnchoredImage,
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  type OneCellAnchor,
  type TwoCellAnchor,
  type WorkbookImage,
  isOneCellAnchor,
  PX_TO_EMU,
} from './core/image.ts';

export {Cell} from './core/cell.ts';
export {
  Worksheet,
  type WorksheetState,
  type WorksheetProperties,
  type ColumnProperties,
  type RowProperties,
  type PageMargins,
  type HeaderFooter,
} from './core/worksheet.ts';
export {
  type SheetProtection,
  type SheetProtectionOptions,
  type SheetProtectionFlags,
  type SheetProtectionCredential,
} from './core/protection.ts';
export {
  PivotTable,
  type PivotMetric,
  type PivotTableOptions,
} from './core/pivot-table.ts';
export {
  Table,
  type TableColumn,
  type TableOptions,
  type TableRegion,
} from './core/table.ts';
export {
  Workbook,
  type WorkbookProperties,
  type AddWorksheetOptions,
} from './core/workbook.ts';

export {writeXlsx} from './io/xlsx/write.ts';
export {readXlsx, type ReadXlsxOptions} from './io/xlsx/read.ts';
export {writeCsv, writeCsvText, type CsvWriteOptions} from './io/csv/write.ts';
export {readCsv, type CsvReadOptions} from './io/csv/read.ts';
export {
  readSheetRows,
  readWorkbookStream,
  type ReadSheetRowsOptions,
  type StreamedCell,
  type StreamedCellStyle,
  type StreamedRow,
  type StreamedSheet,
} from './io/xlsx/read-rows.ts';
