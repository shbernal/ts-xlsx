// Native `customUI` ribbon read: project a typed, read-only view of the ribbon tree (tabs → groups →
// controls + callback names) over the `customUI.xml` / `customUI14.xml` parts a macro workbook carries.
// Those parts are preserved byte-for-byte on round-trip (see `src/core/preserved.ts`); this is a
// projection over them, the same posture as the VBA read view (`src/vba`). Authoring/editing the ribbon
// is out of scope — round-trip fidelity already comes from verbatim preservation.

export {CustomUiParseError} from './errors.ts';
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
} from './ribbon.ts';
