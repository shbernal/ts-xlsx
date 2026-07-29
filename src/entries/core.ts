// Entry barrel for `@shbernal/ts-xlsx/core` — the document model, with no serialisation attached.
//
// This is what a consumer imports to build or inspect a workbook in memory: the classes, the
// property shapes, and the value vocabulary. It pulls in no ZIP, no XML writer and no BIFF12
// decoder. Errors are not here — the whole failure taxonomy lives behind `/errors`, so that a
// caller who only needs to branch on a failure never loads a codec to get the class.

export {
  type CellAddress,
  columnToNumber,
  decodeAddress,
  decodeRange,
  encodeAddress,
  MAX_COLUMN,
  MAX_ROW,
  numberToColumn,
  type RangeAddress,
} from '../core/address.ts';
export type {
  AutoFilter,
  CustomFilter,
  CustomFilterOperator,
  CustomFilterPredicate,
  FilterColumn,
  FilterCriteria,
  ValuesFilter,
} from '../core/autofilter.ts';
export {Cell} from '../core/cell.ts';
export {
  applyTint,
  type ColorResolutionContext,
  DEFAULT_INDEXED_COLORS,
  resolveColor,
  SYSTEM_INDEXED_COLORS,
} from '../core/color-resolution.ts';
export {Column} from '../core/column.ts';
export type {Comment, CommentThread, Mention, Person} from '../core/comment-thread.ts';
export type {
  CfValueObject,
  ConditionalFormatting,
  ConditionalFormattingRule,
} from '../core/conditional-formatting.ts';
export type {
  DataValidation,
  DataValidationEntry,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationType,
} from '../core/data-validation.ts';
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
} from '../core/image.ts';
export type {
  HeaderFooter,
  PageBreak,
  PageMargins,
  PageSetup,
  PrintOptions,
} from '../core/page-setup.ts';
export {
  type ParsedPivotField,
  type ParsedPivotSource,
  type ParsedPivotTable,
  type PivotCacheField,
  type PivotItem,
  type PivotMetric,
  type PivotNumericSummary,
  type PivotRecordCell,
  type PivotSourceKind,
  PivotTable,
  type PivotTableOptions,
} from '../core/pivot-table.ts';
export type {
  PreservedPart,
  PreservedRelationship,
  PreservedRootReference,
  PreservedWorksheetReference,
} from '../core/preserved.ts';
export type {
  SheetProtection,
  SheetProtectionCredential,
  SheetProtectionFlags,
  SheetProtectionOptions,
} from '../core/protection.ts';
export {Range} from '../core/range.ts';
export {Row} from '../core/row.ts';
export type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  CellStyle,
  Color,
  DifferentialStyle,
  Fill,
  FillPatternType,
  Font,
  FontScheme,
  FontVerticalAlignment,
  GradientFill,
  GradientStop,
  HorizontalAlignment,
  NamedCellStyle,
  PatternFill,
  Protection,
  TableStyleNamespace,
  TableStyleTable,
  UnderlineStyle,
  VerticalAlignment,
} from '../core/style.ts';
export {
  Table,
  type TableColumn,
  type TableColumnStyle,
  type TableOptions,
  type TableRegion,
  type TableStyleInfo,
} from '../core/table.ts';
export {
  isTableStyleElementType,
  STRIPE_ELEMENT_TYPES,
  TABLE_STYLE_ELEMENT_TYPES,
  type TableStyle,
  type TableStyleElement,
  type TableStyleElementType,
} from '../core/table-style.ts';
export {
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  parseThemeColorScheme,
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from '../core/theme.ts';
export {
  type CellValue,
  coerceCellValue,
  type DataTableFormulaValue,
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
} from '../core/value.ts';
export {
  type AddImageOptions,
  type AddWorksheetOptions,
  DEFAULT_WORKBOOK_VIEW,
  type DefinedName,
  type PreservedWorkbookReference,
  Workbook,
  type WorkbookProperties,
  type WorkbookView,
} from '../core/workbook.ts';
export type {
  WorkbookProtection,
  WorkbookProtectionCredentialAttr,
} from '../core/workbook-protection.ts';
export {
  type CellModel,
  type ColumnProperties,
  type OutlineProperties,
  type RowInput,
  type RowProperties,
  type SheetView,
  Worksheet,
  type WorksheetModel,
  type WorksheetProperties,
  type WorksheetState,
} from '../core/worksheet.ts';
