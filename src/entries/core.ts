// Entry barrel for `@shbernal/ts-xlsx/core`: the document model, with no serialisation attached.
//
// This is what a consumer imports to build or inspect a workbook in memory: the classes, the
// property shapes, and the value vocabulary. It pulls in no ZIP, no XML writer and no BIFF12
// decoder. Errors are not here: the whole failure taxonomy lives behind `/errors`, so that a
// caller who only needs to branch on a failure never loads a codec to get the class.

export {
  type CellAddress,
  columnToNumber,
  decodeAddress,
  decodeRange,
  encodeAddress,
  type GridRect,
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
export type {Comment, CommentThread, Mention, MentionRef, Person} from '../core/comment-thread.ts';
export type {
  CfTimePeriod,
  CfValueObject,
  CfValueObjectType,
  ConditionalFormatting,
  ConditionalFormattingOperator,
  ConditionalFormattingRule,
  ConditionalFormattingType,
  IconSetType,
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
  type PortableImage,
  PX_TO_EMU,
  type TwoCellAnchor,
  type WorkbookImage,
  type WorksheetImages,
} from '../core/image.ts';
// The OPC container's inflate bound, shared by every reader (`readXlsx`, `readXlsb`, the row
// streamer). It is here rather than on a codec subpath so a `/xlsb` or `/csv` consumer can name it
// without importing the XML codec, and because a container bound is a fact about the package rather
// than about either serialisation inside it.
export type {ReadPackageOptions} from '../io/opc/read-options.ts';
export {
  INVALID_SHEET_NAME_CHARS,
  MAX_COLUMN_WIDTH,
  MAX_ROW_HEIGHT,
  MAX_SHEET_NAME_LENGTH,
  MAX_TABLE_NAME_LENGTH,
  TABLE_NAME_PATTERN,
} from '../core/limits.ts';
export type {
  HeaderFooter,
  PageBreak,
  PageMargins,
  PageOrder,
  PageOrientation,
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
  Fill,
  FillPatternType,
  Font,
  FontScheme,
  FontVerticalAlignment,
  GradientFill,
  GradientStop,
  HorizontalAlignment,
  PatternFill,
  Protection,
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
// The workbook style *tables*: shapes that compose `CellStyle` without being one, so they live beside
// the slice that owns them rather than beside `Fill`. Only the declaration moved; the public surface
// is unchanged.
export type {
  DifferentialStyle,
  NamedCellStyle,
  TableStyleNamespace,
  TableStyleTable,
} from '../core/workbook-styles.ts';
export {
  isTableStyleElementType,
  STRIPE_ELEMENT_TYPES,
  TABLE_STYLE_ELEMENT_TYPES,
  type TableStyle,
  type TableStyleElement,
  type TableStyleElementType,
} from '../core/table-style.ts';
export {estimateWrappedLines} from '../core/text-metrics.ts';
export {
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from '../core/theme.ts';
export {
  type CellValue,
  cellValueToText,
  coerceCellValue,
  type DataTableFormulaValue,
  detectValueType,
  ERROR_CODES,
  type ErrorCode,
  type ErrorValue,
  type FormulaResult,
  type FormulaValue,
  type HyperlinkValue,
  isDataTableFormulaValue,
  isErrorCode,
  isErrorValue,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
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
  type Visibility,
  Worksheet,
  type WorksheetModel,
  type WorksheetProperties,
  type WorksheetState,
} from '../core/worksheet.ts';
