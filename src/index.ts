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

export {type Color, type Font, type UnderlineStyle, type VerticalAlignment} from './core/style.ts';

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
  Workbook,
  type WorkbookProperties,
  type AddWorksheetOptions,
} from './core/workbook.ts';

export {writeXlsx} from './io/xlsx/write.ts';
