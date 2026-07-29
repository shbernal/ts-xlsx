// Entry barrel for `@shbernal/ts-xlsx/customui` — the read-only ribbon view over a macro
// workbook's preserved `customUI` parts.

export {
  CUSTOMUI_2006_NAMESPACE,
  CUSTOMUI_2007_REL_TYPE,
  CUSTOMUI_2009_NAMESPACE,
  CUSTOMUI_2010_REL_TYPE,
  type CustomUiDocument,
  isCustomUiRelType,
  parseCustomUi,
  type Ribbon,
  type RibbonControl,
  type RibbonControlKind,
  type RibbonDialect,
  type RibbonGroup,
  type RibbonTab,
} from '../customui/ribbon.ts';
