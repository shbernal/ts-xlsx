// Worksheet serialisation: a Worksheet model into its `xl/worksheets/sheetN.xml` part (and the sheet's
// rels part and table parts). Owns the row/cell renderer the streaming writer also drives, the sheet's
// print/page/view/protection blocks, and the shared-formula planning that resolves each formula cell's
// role before the row loop.

import {decodeRange, encodeAddress, MAX_COLUMN} from '../../core/address.ts';
import type {AutoFilter, FilterColumn, FilterCriteria} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import {DEFAULT_DATE_NUMFMT, dateToSerial} from '../../core/date.ts';
import {mangleFormula} from '../../core/formula.ts';
import {SHEET_PROTECTION_FLAGS, type SheetProtection} from '../../core/protection.ts';
import type {Table, TableColumn, TableStyleInfo} from '../../core/table.ts';
import {
  detectValueType,
  type FormulaResult,
  isDataTableFormulaValue,
  isErrorValue,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
} from '../../core/value.ts';
import type {
  ColumnProperties,
  HeaderFooter,
  OutlineProperties,
  PageBreak,
  PageMargins,
  PageSetup,
  PrintOptions,
  RowProperties,
  SheetView,
  Worksheet,
  WorksheetProperties,
} from '../../core/worksheet.ts';
import {conditionalFormattingsExtXml, conditionalFormattingsXml} from './conditional-formatting.ts';
import {dataValidationsExtXml, dataValidationsXml} from './data-validation.ts';
import {hyperlinksXml, type PlannedHyperlink} from './hyperlinks.ts';
import {SLICER_LIST_EXT_URI, X14_NS} from './namespaces.ts';
import type {
  BackgroundPlan,
  CommentPlan,
  DrawingPlan,
  PivotPlan,
  PlannedTable,
  PreservedReferencePlan,
  PrinterSettingsPlan,
} from './package-plan.ts';
import {numberText, relativePartPath} from './part-paths.ts';
import {NS, REL, relationship} from './relationships.ts';
import {richTextRunsXml} from './rich-text.ts';
import type {SharedStringTable} from './shared-strings.ts';
import {colorAttrs, type StyleRegistry} from './styles.ts';
import {escapeAttr, escapeText, textElement, XML_DECLARATION} from './xml.ts';

/**
 * A worksheet's eagerly-serialised rows: each row's `<row>` XML tagged with its number (so it merges
 * into ascending order with the sheet's remaining live rows, whatever order it was committed in), plus
 * the used-cell extent they span (or the `Infinity`/`-Infinity` sentinels when the flushed rows carried
 * only formatting). The buffered pass folds the extent into the sheet's dimension.
 */
export interface FlushedSheet {
  readonly rows: ReadonlyArray<{readonly number: number; readonly xml: string}>;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export function worksheetXml(
  sheet: Worksheet,
  tables: readonly PlannedTable[],
  styles: StyleRegistry,
  drawingRelId: string | null,
  legacyDrawingRelId: string | null,
  printerSettingsRelId: string | null,
  backgroundRelId: string | null,
  legacyDrawingHFRelId: string | null,
  preservedSlicerRelIds: readonly string[],
  hyperlinks: readonly PlannedHyperlink[],
  sharedStrings: SharedStringTable | null,
  flushed?: FlushedSheet,
): string {
  // A merge overlapping a table is Excel-invalid geometry; reject it before serialising
  // rather than emit a package a consumer repairs on open.
  validateMerges(sheet);

  const columnDefaults = buildColumnDefaults(sheet);

  // A cell filled from a shared formula is written as a master (seeding the group) or a clone
  // (referencing it by shared index); resolve every such role before the row loop so each cell knows
  // how to serialise its `<f>`. This also validates the master/clone geometry, throwing if a clone
  // precedes its master or its master carries no formula.
  const sharedRoles = planSharedFormulas(sheet);

  // A fully-hidden outline group's collapse toggle belongs on its summary row; derive that set once
  // so the row loop can stamp it even onto a summary row that carries no properties of its own.
  const collapsedSummaries = collapsedSummaryRows(sheet);

  const context: RowRenderContext = {
    columnDefaults,
    styles,
    sharedStrings,
    sharedRoles,
    collapsedSummaries,
  };

  const liveRows: {number: number; xml: string}[] = [];
  // Seed the used-cell extent with any rows the streaming writer already serialised and evicted, so
  // the dimension spans both them and the live rows below.
  let top = flushed?.top ?? Infinity;
  let left = flushed?.left ?? Infinity;
  let bottom = flushed?.bottom ?? -Infinity;
  let right = flushed?.right ?? -Infinity;

  for (const entry of sheet.rows()) {
    const {xml, minCol, maxCol} = renderRow(entry, context);
    if (xml === '') continue;
    liveRows.push({number: entry.number, xml});
    if (minCol !== Infinity) {
      if (entry.number < top) top = entry.number;
      if (entry.number > bottom) bottom = entry.number;
      if (minCol < left) left = minCol;
      if (maxCol > right) right = maxCol;
    }
  }

  // Dimension is the used *cell* range; rows/columns carrying only formatting do not
  // extend it, matching how Excel records <dimension>.
  const dimensionRef =
    bottom === -Infinity ? 'A1' : `${encodeAddress(left, top)}:${encodeAddress(right, bottom)}`;
  // Merge the streaming writer's pre-rendered rows with the live ones into ascending row order — a
  // flushed row can carry any number, and rows may be committed out of order. The buffered path has no
  // flushed rows, so it skips the merge and its sort entirely.
  const orderedRows = flushed
    ? [...flushed.rows, ...liveRows].sort((a, b) => a.number - b.number)
    : liveRows;
  const bodyXml = orderedRows.map((row) => row.xml).join('');
  const sheetData = bodyXml === '' ? '<sheetData/>' : `<sheetData>${bodyXml}</sheetData>`;

  return (
    XML_DECLARATION +
    `<worksheet xmlns="${NS.main}" xmlns:r="${NS.docRels}">` +
    sheetPrXml(sheet) +
    `<dimension ref="${dimensionRef}"/>` +
    sheetViewsXml(sheet.view) +
    sheetFormatPr(sheet.properties) +
    colsXml(sheet, styles) +
    sheetData +
    sheetProtectionXml(sheet.protection) +
    // CT_Worksheet order: <autoFilter> follows <sheetProtection> (and the scenarios block) and
    // precedes <mergeCells>. Its `_FilterDatabase` companion is emitted in the workbook part.
    autoFilterXml(sheet.autoFilter) +
    mergeCellsXml(sheet.merges) +
    // CT_Worksheet order: <conditionalFormatting> blocks follow <mergeCells>, then <dataValidations>,
    // then <hyperlinks> — all precede the print settings.
    conditionalFormattingsXml(sheet.conditionalFormattings, styles) +
    dataValidationsXml(sheet.dataValidations) +
    hyperlinksXml(hyperlinks) +
    // CT_Worksheet order: <printOptions> precedes <pageMargins>, which precedes <pageSetup>.
    printOptionsXml(sheet.printOptions) +
    pageMarginsXml(sheet.pageMargins) +
    pageSetupXml(sheet.pageSetup, printerSettingsRelId) +
    headerFooterXml(sheet.headerFooter) +
    // CT_Worksheet order: <rowBreaks> follows <headerFooter>, <colBreaks> follows <rowBreaks>, and
    // both precede the drawing block.
    pageBreaksXml(sheet.rowBreaks, 'rowBreaks') +
    pageBreaksXml(sheet.columnBreaks, 'colBreaks') +
    // Schema order near the tail: <drawing> (the images), then <legacyDrawing> (the VML holding the
    // note boxes), then <legacyDrawingHF> (a preserved header/footer image's VML), then <picture>
    // (the sheet background), then <tableParts>.
    (drawingRelId !== null ? `<drawing r:id="${drawingRelId}"/>` : '') +
    (legacyDrawingRelId !== null ? `<legacyDrawing r:id="${legacyDrawingRelId}"/>` : '') +
    (legacyDrawingHFRelId !== null ? `<legacyDrawingHF r:id="${legacyDrawingHFRelId}"/>` : '') +
    (backgroundRelId !== null ? `<picture r:id="${backgroundRelId}"/>` : '') +
    tablePartsXml(tables) +
    // `<extLst>` is the final child of CT_Worksheet and a worksheet may carry at most one. Both the
    // x14 conditional-formatting extensions (data-bar gradient/negative-fill/axis) and the extended
    // (x14) data validations ride inside it as sibling `<ext>` blocks — so they are gathered here into
    // a single `<extLst>` rather than each emitting its own.
    worksheetExtLstXml(sheet, preservedSlicerRelIds) +
    '</worksheet>'
  );
}

/**
 * A column's style facets are defaults its cells inherit unless they override them; the writer
 * composes each cell's full style up front (cell over row over column, per facet) so a cell that
 * overrides one facet still carries the column's others, rather than silently dropping them. Frozen
 * once by the streaming writer at its first flush so every eagerly-rendered row sees the same defaults.
 */
export function buildColumnDefaults(sheet: Worksheet): Map<number, ColumnProperties> {
  const columnDefaults = new Map<number, ColumnProperties>();
  for (const {index, properties} of sheet.columns()) columnDefaults.set(index, properties);
  return columnDefaults;
}

/** The whole-sheet context a single row needs to serialise: the column defaults it inherits, the
 * style/string tables it interns into, the shared-formula roles its cells play, and the collapsed
 * outline summaries whose toggle it must stamp. The streaming writer supplies empty shared-formula
 * and collapsed-summary sets, since those are whole-sheet derivations a flushed row cannot join. */
export interface RowRenderContext {
  readonly columnDefaults: ReadonlyMap<number, ColumnProperties>;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly sharedRoles: ReadonlyMap<string, SharedFormulaRole>;
  readonly collapsedSummaries: ReadonlySet<number>;
}

/**
 * Serialise one row to its `<row>` element, or '' when the row has neither data nor its own
 * formatting. Returns the used-column bounds (`Infinity`/`-Infinity` when nothing was rendered) so a
 * caller can fold them into the sheet dimension. Shared by the buffered sheet pass and the streaming
 * writer's eager flush, so both emit byte-identical rows.
 */
export function renderRow(
  entry: {
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
  },
  ctx: RowRenderContext,
): {xml: string; minCol: number; maxCol: number} {
  const {number, cells, properties} = entry;
  // A cell earns a <c> element if it holds a value OR carries its own style: a formatted-but-empty
  // cell (a fill/border on a null value) is a real cell to Excel, and dropping it would lose the
  // formatting. A cell with neither is inherited from its row/column and needs no element of its own.
  const rendered = cells.filter((cell) => cell.value !== null || hasOwnStyle(cell));
  const attrs = rowAttrs(properties, ctx.styles, ctx.collapsedSummaries.has(number));
  // A row with neither data nor its own formatting has nothing to serialise.
  if (rendered.length === 0 && attrs === '') return {xml: '', minCol: Infinity, maxCol: -Infinity};
  const rowFill = properties?.fill;
  const cellsXml = rendered
    .map((cell) => {
      // Cell overrides win over row/column defaults; a cell with any facet gets its own,
      // fully-composed style entry so no default facet is lost to the override. Precedence is
      // cell over row over column per facet — the row contributes only a fill today.
      const colDef = ctx.columnDefaults.get(cell.col);
      const style = ctx.styles.styleId({
        fill: cell.fill ?? rowFill ?? colDef?.fill,
        // A bare Date carries no format of its own, so it renders as a raw serial and reads
        // back as a number unless we apply a date format. An explicit cell/column format wins.
        numFmt: cell.numFmt ?? colDef?.numFmt ?? dateDefaultNumFmt(cell.value),
        font: cell.font ?? colDef?.font,
        border: cell.border ?? colDef?.border,
        alignment: cell.alignment ?? colDef?.alignment,
        protection: cell.protection ?? colDef?.protection,
        // Quote-prefix is a cell-only flag; a column carries no such default to inherit.
        quotePrefix: cell.quotePrefix,
        // The cell's link into the named cell-style layer, preserved so a round-trip keeps it tied
        // to that style rather than flattening it into a purely-direct format.
        xfId: cell.namedStyleId,
      });
      return cellXml(cell, style, ctx.sharedRoles.get(cell.address), ctx.sharedStrings);
    })
    .join('');
  let minCol = Infinity;
  let maxCol = -Infinity;
  for (const cell of rendered) {
    if (cell.col < minCol) minCol = cell.col;
    if (cell.col > maxCol) maxCol = cell.col;
  }
  return {xml: `<row r="${number}"${attrs}>${cellsXml}</row>`, minCol, maxCol};
}

// Assemble the worksheet's single `<extLst>` from every x14 extension the sheet carries, or '' when it
// carries none. Each producer returns a bare `<ext>` so they compose without nesting an `<extLst>`.
function worksheetExtLstXml(sheet: Worksheet, slicerRelIds: readonly string[]): string {
  const exts = [
    conditionalFormattingsExtXml(sheet.conditionalFormattings),
    dataValidationsExtXml(sheet.dataValidations),
    slicerListExtXml(slicerRelIds),
  ].filter((ext) => ext !== '');
  return exts.length === 0 ? '' : `<extLst>${exts.join('')}</extLst>`;
}

// The worksheet-body `<x14:slicerList>` extension that reconnects a sheet to its preserved slicer
// parts. Each `<x14:slicer>` names the sheet-local relationship id its slicer rel was re-emitted under,
// so the wiring stays consistent even though the id is reassigned on write. '' when the sheet has none.
function slicerListExtXml(slicerRelIds: readonly string[]): string {
  if (slicerRelIds.length === 0) return '';
  const slicers = slicerRelIds.map((relId) => `<x14:slicer r:id="${relId}"/>`).join('');
  return (
    `<ext uri="${SLICER_LIST_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:slicerList>${slicers}</x14:slicerList></ext>`
  );
}

// Excel forbids a merged range from intersecting a formatted table; such a file opens as
// corrupt. The writer is the OOXML gatekeeper for this cross-feature geometry conflict.
function validateMerges(sheet: Worksheet): void {
  if (sheet.merges.length === 0 || sheet.tables.length === 0) return;
  for (const merge of sheet.merges) {
    const {left, right, top, bottom} = decodeRange(merge);
    if (left === undefined || right === undefined || top === undefined || bottom === undefined)
      continue;
    for (const table of sheet.tables) {
      const region = table.region;
      const overlaps =
        left <= region.right &&
        right >= region.left &&
        top <= region.bottom &&
        bottom >= region.top;
      if (overlaps) {
        throw new Error(
          `merged range ${merge} overlaps table "${table.name}" (${table.ref}) — Excel forbids a merge inside a table`,
        );
      }
    }
  }
}

// `<sheetViews>` holds the sheet's single view. A frozen view adds a `<pane>` recording the split
// and a `<selection>` naming the pane the split activates, exactly as Excel writes it — a normal
// view carries neither, so unfreezing leaves no leftover `<pane>` that would trip a repair prompt.
// The active pane is whichever scrolling region the freeze creates: bottom-right when both axes are
// frozen, else top-right (columns only) or bottom-left (rows only).
function sheetViewsXml(view: SheetView): string {
  const xSplit = view.xSplit ?? 0;
  const ySplit = view.ySplit ?? 0;
  if (view.state !== 'frozen' || (xSplit === 0 && ySplit === 0)) {
    return '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
  }
  const topLeftCell = view.topLeftCell ?? encodeAddress(xSplit + 1, ySplit + 1);
  const activePane =
    xSplit > 0 && ySplit > 0 ? 'bottomRight' : xSplit > 0 ? 'topRight' : 'bottomLeft';
  const pane =
    '<pane' +
    (xSplit > 0 ? ` xSplit="${xSplit}"` : '') +
    (ySplit > 0 ? ` ySplit="${ySplit}"` : '') +
    ` topLeftCell="${escapeAttr(topLeftCell)}" activePane="${activePane}" state="frozen"/>`;
  const selection = `<selection pane="${activePane}" activeCell="${escapeAttr(topLeftCell)}" sqref="${escapeAttr(topLeftCell)}"/>`;
  return `<sheetViews><sheetView workbookViewId="0">${pane}${selection}</sheetView></sheetViews>`;
}

// `<sheetPr>` carries the sheet's appearance properties: the tab colour, the outline
// summary-position flags, and the fit-to-page flag. It is the first child of `<worksheet>` in
// CT_Worksheet order; its own children follow CT_SheetPr order — `<tabColor>`, `<outlinePr>`, then
// `<pageSetUpPr>`. Omitted entirely when the sheet carries none, so an unadorned sheet stays
// byte-clean.
function sheetPrXml(sheet: Worksheet): string {
  const children =
    (sheet.tabColor !== undefined ? `<tabColor ${colorAttrs(sheet.tabColor)}/>` : '') +
    outlinePrXml(sheet.outline) +
    pageSetUpPrXml(sheet.pageSetup);
  return children === '' ? '' : `<sheetPr>${children}</sheetPr>`;
}

// `<pageSetUpPr>` holds the fit-to-page toggle, which lives on the sheet properties rather than on
// `<pageSetup>` — Excel reads it from here to decide whether the `fitToWidth`/`fitToHeight` counts
// or the fixed `scale` govern printing. Emitted only when the author set the flag.
function pageSetUpPrXml(pageSetup: PageSetup): string {
  return pageSetup.fitToPage ? '<pageSetUpPr fitToPage="1"/>' : '';
}

// `<outlinePr>` carries only the summary-position flags today. Each is emitted solely when the
// caller set it, so an inverted placement (`summaryBelow="0"`) is honoured while an untouched sheet
// keeps the element out of the file entirely.
function outlinePrXml(outline: OutlineProperties): string {
  const attrs: string[] = [];
  if (outline.summaryBelow !== undefined)
    attrs.push(`summaryBelow="${outline.summaryBelow ? 1 : 0}"`);
  if (outline.summaryRight !== undefined)
    attrs.push(`summaryRight="${outline.summaryRight ? 1 : 0}"`);
  return attrs.length === 0 ? '' : `<outlinePr ${attrs.join(' ')}/>`;
}

function mergeCellsXml(merges: readonly string[]): string {
  if (merges.length === 0) return '';
  const cells = merges
    .map((range) => `<mergeCell ref="${escapeAttr(decodeRange(range).dimensions)}"/>`)
    .join('');
  return `<mergeCells count="${merges.length}">${cells}</mergeCells>`;
}

// The sheet's autofilter: `<autoFilter ref="A1:C10"/>` when it only draws dropdowns, or with nested
// `<filterColumn>` children when columns carry criteria. Its companion `_FilterDatabase` defined name
// (the range Excel derives filtering from) is written in the workbook part, so a sheet with no filter
// emits nothing here and nothing there.
function autoFilterXml(filter: AutoFilter | undefined): string {
  if (filter === undefined) return '';
  const ref = escapeAttr(filter.ref);
  if (filter.columns.length === 0) return `<autoFilter ref="${ref}"/>`;
  return `<autoFilter ref="${ref}">${filter.columns.map(filterColumnXml).join('')}</autoFilter>`;
}

function filterColumnXml(column: FilterColumn): string {
  return `<filterColumn colId="${column.colId}">${filterCriteriaXml(column.criteria)}</filterColumn>`;
}

// A values filter is `<filters>` with a `<filter val>` per allowed value (and `blank="1"` to admit
// empty cells); a custom filter is `<customFilters>` with one or two `<customFilter operator val>`
// predicates, `and="1"` when they are AND-combined rather than OR.
function filterCriteriaXml(criteria: FilterCriteria): string {
  if (criteria.kind === 'values') {
    const blankAttr = criteria.blank ? ' blank="1"' : '';
    const filters = criteria.values.map((value) => `<filter val="${escapeAttr(value)}"/>`).join('');
    return `<filters${blankAttr}>${filters}</filters>`;
  }
  const andAttr = criteria.and ? ' and="1"' : '';
  const predicates = criteria.predicates
    .map((p) => `<customFilter operator="${p.operator}" val="${escapeAttr(p.val)}"/>`)
    .join('');
  return `<customFilters${andAttr}>${predicates}</customFilters>`;
}

// Each sheet-protection flag maps to a `<sheetProtection>` attribute whose value is INVERTED
// from the author-facing allow-flag: the attribute records that an operation is *forbidden*
// ("1"), so `allow: true` serialises as "0". Only a value that differs from OOXML's per-
// attribute default (see SHEET_PROTECTION_FLAGS) is written — most editing operations default
// to forbidden under protection, while selecting cells defaults to permitted.
//
// <sheetProtection> is what makes the per-cell locked/hidden flags bite. `sheet="1"` marks the
// sheet protected; the password credential (when present) guards lifting it; the flag attributes
// carve out the operations that stay available. base64 salt/hash use only XML-safe characters.
function sheetProtectionXml(protection: SheetProtection | undefined): string {
  if (protection === undefined) return '';
  const {flags, credential} = protection;
  let attrs = '';
  if (credential !== undefined) {
    attrs +=
      ` algorithmName="${credential.algorithmName}"` +
      ` hashValue="${credential.hashValue}"` +
      ` saltValue="${credential.saltValue}"` +
      ` spinCount="${credential.spinCount}"`;
  }
  attrs += ' sheet="1"';
  for (const {key, defaultForbidden} of SHEET_PROTECTION_FLAGS) {
    const allow = flags[key];
    if (allow === undefined) continue;
    const forbidden = !allow;
    if (forbidden === defaultForbidden) continue;
    attrs += ` ${key}="${forbidden ? 1 : 0}"`;
  }
  return `<sheetProtection${attrs}/>`;
}

function tablePartsXml(tables: readonly PlannedTable[]): string {
  if (tables.length === 0) return '';
  const parts = tables.map(({relId}) => `<tablePart r:id="${relId}"/>`).join('');
  return `<tableParts count="${tables.length}">${parts}</tableParts>`;
}

export function worksheetRelsXml(
  tables: readonly PlannedTable[],
  drawing: DrawingPlan | null,
  comments: CommentPlan | null,
  printerSettings: PrinterSettingsPlan | null,
  background: BackgroundPlan | null,
  hyperlinks: readonly PlannedHyperlink[],
  preservedReferences: readonly PreservedReferencePlan[],
  pivots: readonly PivotPlan[],
): string {
  const rels = [
    ...tables.map(({relId, number}) =>
      relationship(relId, REL.table, `../tables/table${number}.xml`),
    ),
    // A pivot table hosted on this sheet is reached by a relationship of type pivotTable; Excel
    // discovers the pivot from the rels part, so the sheet body itself carries no reference to it.
    ...pivots.map((pivot) =>
      relationship(
        pivot.sheetRelId,
        REL.pivotTable,
        `../pivotTables/pivotTable${pivot.number}.xml`,
      ),
    ),
    ...(drawing === null
      ? []
      : [relationship(drawing.relId, REL.drawing, `../drawings/drawing${drawing.number}.xml`)]),
    ...(comments === null
      ? []
      : [
          relationship(
            comments.vmlRelId,
            REL.vmlDrawing,
            `../drawings/vmlDrawing${comments.number}.vml`,
          ),
          relationship(comments.commentsRelId, REL.comments, `../comments${comments.number}.xml`),
        ]),
    ...(printerSettings === null
      ? []
      : [
          relationship(
            printerSettings.relId,
            REL.printerSettings,
            `../printerSettings/printerSettings${printerSettings.number}.bin`,
          ),
        ]),
    ...(background === null
      ? []
      : [
          relationship(
            background.relId,
            REL.image,
            `../media/image${background.mediaNumber}.${background.extension}`,
          ),
        ]),
    // A preserved reference targets its entry part's new (package-absolute) path; a worksheet always
    // lives under `xl/worksheets/`, so the target is that path made relative to that directory.
    ...preservedReferences.map((reference) =>
      relationship(
        reference.relId,
        reference.relType,
        escapeAttr(relativePartPath('xl/worksheets/sheet1.xml', reference.entryPath)),
      ),
    ),
    // An external hyperlink's target is a URL outside the package, so its relationship carries
    // TargetMode="External" and the plain `relationship()` helper (a package-internal target) will
    // not do. Internal links have no relId and contribute nothing here.
    ...hyperlinks
      .filter((link) => link.relId !== undefined && link.target !== undefined)
      .map(
        (link) =>
          `<Relationship Id="${link.relId}" Type="${REL.hyperlink}" ` +
          `Target="${escapeAttr(link.target as string)}" TargetMode="External"/>`,
      ),
  ].join('');
  return `${XML_DECLARATION}<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

export function tableXml(table: Table, id: number): string {
  const name = escapeAttr(table.name);
  const displayName = escapeAttr(table.displayName);
  // headerRowCount defaults to 1 in OOXML, so only a headerless table needs it stated.
  const headerRowCount = table.headerRow ? '' : ' headerRowCount="0"';
  // A present totals row implies it is shown, so it only needs the count. Without a totals row the
  // model's tri-state totalsRowShown decides: emit the flag Excel recorded, or nothing when the
  // source omitted it — injecting `totalsRowShown="0"` onto a table that lacked the attribute is
  // exactly the spurious change that makes Excel treat an otherwise-valid table as corrupt.
  let totals: string;
  if (table.totalsRow) {
    totals = ' totalsRowCount="1"';
  } else if (table.totalsRowShown !== undefined) {
    totals = ` totalsRowShown="${table.totalsRowShown ? '1' : '0'}"`;
  } else {
    totals = '';
  }
  const autoFilter =
    table.autoFilterRef !== null ? `<autoFilter ref="${table.autoFilterRef}"/>` : '';
  const columns = table.columns.map((column, i) => tableColumnXml(column, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<table xmlns="${NS.main}" id="${id}" name="${name}" displayName="${displayName}" ` +
    `ref="${table.ref}"${headerRowCount}${totals}>` +
    autoFilter +
    `<tableColumns count="${table.columns.length}">${columns}</tableColumns>` +
    tableStyleInfoXml(table.style) +
    '</table>'
  );
}

// Excel's default table appearance, written for a table that carries no style of its own.
const DEFAULT_TABLE_STYLE =
  '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" ' +
  'showRowStripes="1" showColumnStripes="0"/>';

// Emit `<tableStyleInfo>` from the model's style, or the default when none was captured. Each
// attribute is written only when the model holds it, so a style read without (say) a `name` — or a
// part that omitted a banding flag — re-emits exactly as it arrived rather than gaining an attribute.
function tableStyleInfoXml(style: TableStyleInfo | undefined): string {
  if (style === undefined) return DEFAULT_TABLE_STYLE;
  let attrs = '';
  if (style.name !== undefined) attrs += ` name="${escapeAttr(style.name)}"`;
  if (style.showFirstColumn !== undefined)
    attrs += ` showFirstColumn="${style.showFirstColumn ? '1' : '0'}"`;
  if (style.showLastColumn !== undefined)
    attrs += ` showLastColumn="${style.showLastColumn ? '1' : '0'}"`;
  if (style.showRowStripes !== undefined)
    attrs += ` showRowStripes="${style.showRowStripes ? '1' : '0'}"`;
  if (style.showColumnStripes !== undefined) {
    attrs += ` showColumnStripes="${style.showColumnStripes ? '1' : '0'}"`;
  }
  return `<tableStyleInfo${attrs}/>`;
}

function tableColumnXml(column: TableColumn, id: number): string {
  let attrs = `id="${id}" name="${escapeAttr(column.name)}"`;
  if (column.totalsRowLabel !== undefined) {
    attrs += ` totalsRowLabel="${escapeAttr(column.totalsRowLabel)}"`;
  }
  if (column.totalsRowFunction !== undefined) {
    attrs += ` totalsRowFunction="${escapeAttr(column.totalsRowFunction)}"`;
  }
  return `<tableColumn ${attrs}/>`;
}

// CT_HeaderFooter child order, paired with the flag their presence gates: the even- and
// first-page variants are silently ignored by Excel unless differentOddEven / differentFirst
// are set, so the writer derives each flag from whether any variant in its class was provided.
const HF_CHILDREN = [
  {tag: 'oddHeader', key: 'oddHeader'},
  {tag: 'oddFooter', key: 'oddFooter'},
  {tag: 'evenHeader', key: 'evenHeader'},
  {tag: 'evenFooter', key: 'evenFooter'},
  {tag: 'firstHeader', key: 'firstHeader'},
  {tag: 'firstFooter', key: 'firstFooter'},
] as const;

function headerFooterXml(hf: HeaderFooter): string {
  const children = HF_CHILDREN.filter(({key}) => hf[key] !== undefined);
  if (children.length === 0) return '';
  const differentOddEven = hf.evenHeader !== undefined || hf.evenFooter !== undefined;
  const differentFirst = hf.firstHeader !== undefined || hf.firstFooter !== undefined;
  let attrs = '';
  if (differentOddEven) attrs += ' differentOddEven="1"';
  if (differentFirst) attrs += ' differentFirst="1"';
  const body = children
    .map(({tag, key}) => `<${tag}>${escapeText(hf[key] as string)}</${tag}>`)
    .join('');
  return `<headerFooter${attrs}>${body}</headerFooter>`;
}

// Excel's "Normal" margins, in inches — the defaults Excel writes for an untouched sheet.
const DEFAULT_MARGINS = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
} as const;
const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

// `<printOptions>` carries the print-toggle flags and sits just before `<pageMargins>` in
// CT_Worksheet order. Each attribute is emitted only when the model carries it — as an explicit
// `="1"`/`="0"` so a caller can force a flag off against Excel's default — and an untouched sheet
// keeps the element out of the file entirely.
function printOptionsXml(printOptions: PrintOptions): string {
  const bit = (value: boolean): '1' | '0' => (value ? '1' : '0');
  const attrs: string[] = [];
  if (printOptions.horizontalCentered !== undefined) {
    attrs.push(`horizontalCentered="${bit(printOptions.horizontalCentered)}"`);
  }
  if (printOptions.verticalCentered !== undefined) {
    attrs.push(`verticalCentered="${bit(printOptions.verticalCentered)}"`);
  }
  if (printOptions.headings !== undefined) attrs.push(`headings="${bit(printOptions.headings)}"`);
  if (printOptions.gridLines !== undefined)
    attrs.push(`gridLines="${bit(printOptions.gridLines)}"`);
  if (printOptions.gridLinesSet !== undefined) {
    attrs.push(`gridLinesSet="${bit(printOptions.gridLinesSet)}"`);
  }
  return attrs.length === 0 ? '' : `<printOptions ${attrs.join(' ')}/>`;
}

// OOXML's <pageMargins> is all-or-nothing: setting any one margin requires all six, or Excel
// repairs the file. So the element is emitted only when the caller set at least one, and the
// untouched sides fall back to the Normal-preset defaults.
function pageMarginsXml(margins: PageMargins): string {
  if (MARGIN_SIDES.every((side) => margins[side] === undefined)) return '';
  const attrs = MARGIN_SIDES.map(
    (side) => `${side}="${numberText(margins[side] ?? DEFAULT_MARGINS[side])}"`,
  ).join(' ');
  return `<pageMargins ${attrs}/>`;
}

// `<pageSetup>` carries the print-scaling attributes (all but `fitToPage`, which is a `<sheetPr>`
// flag). It sits between `<pageMargins>` and `<headerFooter>` in CT_Worksheet order. Each attribute
// is emitted only when the author set it, so an untouched sheet keeps the element out of the file
// and a partially-set one never fabricates the counts Excel would otherwise default. A non-null
// `printerSettingsRelId` links the sheet's opaque printer-settings blob and forces the element out
// even when no scaling attribute is set — the reference is the only thing the model has to carry.
function pageSetupXml(pageSetup: PageSetup, printerSettingsRelId: string | null): string {
  const attrs: string[] = [];
  if (pageSetup.paperSize !== undefined) attrs.push(`paperSize="${pageSetup.paperSize}"`);
  if (pageSetup.scale !== undefined) attrs.push(`scale="${pageSetup.scale}"`);
  if (pageSetup.fitToWidth !== undefined) attrs.push(`fitToWidth="${pageSetup.fitToWidth}"`);
  if (pageSetup.fitToHeight !== undefined) attrs.push(`fitToHeight="${pageSetup.fitToHeight}"`);
  if (pageSetup.pageOrder !== undefined) attrs.push(`pageOrder="${pageSetup.pageOrder}"`);
  if (pageSetup.orientation !== undefined) attrs.push(`orientation="${pageSetup.orientation}"`);
  if (printerSettingsRelId !== null) attrs.push(`r:id="${printerSettingsRelId}"`);
  return attrs.length === 0 ? '' : `<pageSetup ${attrs.join(' ')}/>`;
}

function sheetFormatPr(properties: WorksheetProperties): string {
  const rowHeight = properties.defaultRowHeight ?? 15;
  let attrs = ` defaultRowHeight="${numberText(rowHeight)}"`;
  if (properties.defaultColWidth !== undefined) {
    attrs += ` defaultColWidth="${numberText(properties.defaultColWidth)}"`;
  }
  // A non-standard default row height is only honoured by Excel when customHeight is set.
  if (properties.defaultRowHeight !== undefined) attrs += ' customHeight="1"';
  return `<sheetFormatPr${attrs}/>`;
}

function colsXml(sheet: Worksheet, styles: StyleRegistry): string {
  // Runs of adjacent columns that carry identical definitions are coalesced into a single
  // `<col min max>` span — Excel writes columns this way, and it keeps the part compact for a
  // sheet whose columns share a width or outline level. A gap in the indices or any difference
  // in the emitted attributes breaks the run.
  const runs: {min: number; max: number; body: string}[] = [];
  for (const {index, properties} of sheet.columns()) {
    // OOXML has no column past XFD (16384); a definition beyond it is corrupt to Excel,
    // so drop it rather than emit an out-of-range <col> range.
    if (index > MAX_COLUMN) continue;
    const body = colBody(properties, styles);
    // A <col> with no width, visibility, or style says nothing; omit it entirely.
    if (body === null) continue;
    const last = runs[runs.length - 1];
    if (last !== undefined && last.max === index - 1 && last.body === body) {
      last.max = index;
    } else {
      runs.push({min: index, max: index, body});
    }
  }
  if (runs.length === 0) return '';
  const cols = runs.map((run) => `<col min="${run.min}" max="${run.max}"${run.body}/>`).join('');
  return `<cols>${cols}</cols>`;
}

// The attributes of a `<col>` sans its `min`/`max` span (each with a leading space), or `null` when
// the column declares nothing worth emitting. Two columns with the same body are interchangeable, so
// the body doubles as the equivalence key that {@link colsXml} coalesces adjacent runs by.
function colBody(properties: ColumnProperties, styles: StyleRegistry): string | null {
  let attrs = '';
  let meaningful = false;
  if (properties.width !== undefined) {
    attrs += ` width="${numberText(properties.width)}" customWidth="1"`;
    meaningful = true;
  }
  if (properties.hidden) {
    attrs += ' hidden="1"';
    meaningful = true;
  }
  if (properties.outlineLevel !== undefined && properties.outlineLevel > 0) {
    attrs += ` outlineLevel="${properties.outlineLevel}"`;
    meaningful = true;
  }
  if (properties.collapsed) {
    attrs += ' collapsed="1"';
    meaningful = true;
  }
  // The column's style facets are carried as its own `<col>` style; its populated cells inherit
  // them via the composition above, and this `style` makes Excel apply them to the column's empty
  // cells too.
  const style = styles.styleId({
    fill: properties.fill,
    numFmt: properties.numFmt,
    font: properties.font,
    border: properties.border,
    alignment: properties.alignment,
    protection: properties.protection,
  });
  if (style !== 0) {
    attrs += ` style="${style}"`;
    meaningful = true;
  }
  return meaningful ? attrs : null;
}

// Manual page breaks (`<rowBreaks>`/`<colBreaks>`): one `<brk>` per row/column the layout splits
// before. Excel records both the running total (`count`) and the manual subset (`manualBreakCount`);
// every break the model carries is a manual, author-set one, so the two counts coincide. `max` bounds
// the break across the other axis (Excel writes the last row/column index); a break without one is
// emitted bare. Row and column breaks share this shape, differing only in the wrapping element.
function pageBreaksXml(breaks: readonly PageBreak[], element: 'rowBreaks' | 'colBreaks'): string {
  if (breaks.length === 0) return '';
  const brks = breaks
    .map((brk) => {
      const maxAttr = brk.max !== undefined ? ` max="${brk.max}"` : '';
      return `<brk id="${brk.id}"${maxAttr} man="1"/>`;
    })
    .join('');
  return `<${element} count="${breaks.length}" manualBreakCount="${breaks.length}">${brks}</${element}>`;
}

function rowAttrs(
  properties: RowProperties | undefined,
  styles: StyleRegistry,
  collapsedSummary: boolean,
): string {
  if (properties === undefined) return collapsedSummary ? ' collapsed="1"' : '';
  let attrs = '';
  if (properties.height !== undefined)
    attrs += ` ht="${numberText(properties.height)}" customHeight="1"`;
  if (properties.hidden) attrs += ' hidden="1"';
  if (properties.outlineLevel !== undefined && properties.outlineLevel > 0) {
    attrs += ` outlineLevel="${properties.outlineLevel}"`;
  }
  // The collapse toggle is set explicitly by the author, or derived onto a summary row whose whole
  // detail group is hidden (see {@link collapsedSummaryRows}). It rides the summary row, never the
  // detail rows.
  if (properties.collapsed || collapsedSummary) attrs += ' collapsed="1"';
  // A row-level fill is a default format for the row's cells; customFormat="1" is what makes
  // Excel honour the row's `s`, and a cell without its own `s` then inherits it.
  const style = styles.styleId({fill: properties.fill});
  if (style !== 0) attrs += ` s="${style}" customFormat="1"`;
  return attrs;
}

// A collapsed outline group is two coordinated facts: its detail rows carry outlineLevel and are
// hidden, AND the summary row that terminates the group carries `collapsed`. Authors typically set
// only outlineLevel + hidden on the detail rows, so the summary flag is derived here rather than
// demanded of the caller: a row is a collapsed summary iff its adjacent detail run — the contiguous
// higher-outline-level rows on the summary side — is non-empty and every row in it is hidden.
// Placement follows the sheet's summaryBelow flag (Excel's default is summary below the detail); the
// walk stops at the first row of level <= the summary's own, so a gap or a boundary ends the group.
function collapsedSummaryRows(sheet: Worksheet): Set<number> {
  const level = new Map<number, number>();
  const hidden = new Map<number, boolean>();
  for (const {number, properties} of sheet.rows()) {
    level.set(number, properties?.outlineLevel ?? 0);
    hidden.set(number, properties?.hidden ?? false);
  }
  const levelOf = (row: number): number => level.get(row) ?? 0;
  const step = sheet.outline.summaryBelow === false ? 1 : -1;
  const summaries = new Set<number>();
  for (const [summary, summaryLevel] of level) {
    let detail = summary + step;
    let sawDetail = false;
    let allHidden = true;
    while (levelOf(detail) > summaryLevel) {
      sawDetail = true;
      if (!hidden.get(detail)) allHidden = false;
      detail += step;
    }
    if (sawDetail && allHidden) summaries.add(summary);
  }
  return summaries;
}

// A valid Date — whether the cell's own value or a formula's cached result — with no format of its
// own gets the default date format so it renders and reads back as a date rather than a bare serial.
// An Invalid Date and every non-date value contribute nothing here.
function dateDefaultNumFmt(value: Cell['value']): string | undefined {
  const date =
    value instanceof Date
      ? value
      : (isFormulaValue(value) || isSharedFormulaValue(value)) && value.result instanceof Date
        ? value.result
        : undefined;
  return date !== undefined && !Number.isNaN(date.getTime()) ? DEFAULT_DATE_NUMFMT : undefined;
}

function cellXml(
  cell: Cell,
  style: number,
  shared: SharedFormulaRole | undefined,
  sharedStrings: SharedStringTable | null,
): string {
  const ref = cell.address;
  const value = cell.value;
  const s = style !== 0 ? ` s="${style}"` : '';

  // A shared-formula master seeds the group with its formula text under `t="shared" ref si`; a clone
  // carries no text of its own, only a back-reference to the master's `si`. Its cached result still
  // travels with the cell.
  if (shared !== undefined) {
    if (shared.ref !== undefined && isFormulaValue(value)) {
      const f = `<f t="shared" ref="${shared.ref}" si="${shared.si}">${escapeText(mangleFormula(value.formula))}</f>`;
      return formulaBodyXml(ref, s, f, value.result);
    }
    const result = isSharedFormulaValue(value) ? value.result : undefined;
    return formulaBodyXml(ref, s, `<f t="shared" si="${shared.si}"/>`, result);
  }

  if (isDataTableFormulaValue(value)) {
    // A data-table formula carries no expression text — only its declaration attributes — which we
    // re-emit verbatim so a read-modify-write cycle preserves the What-If kind the library never
    // evaluates. The cached result travels as any formula result does.
    const attrs =
      `ref="${escapeAttr(value.ref)}"` +
      ` dt2D="${value.dataTable2D ? 1 : 0}"` +
      ` dtr="${value.dataTableRow ? 1 : 0}"` +
      (value.r1 !== undefined ? ` r1="${escapeAttr(value.r1)}"` : '') +
      (value.r2 !== undefined ? ` r2="${escapeAttr(value.r2)}"` : '');
    return formulaBodyXml(ref, s, `<f t="dataTable" ${attrs}/>`, value.result);
  }
  if (isFormulaValue(value)) {
    return formulaCellXml(ref, s, value.formula, value.result);
  }
  if (value instanceof Date) {
    // An Invalid Date (new Date(NaN)) has no serial; keep the cell (and its style) but emit no
    // value rather than throwing, so one bad date never takes down the whole sheet's export.
    if (Number.isNaN(value.getTime())) return `<c r="${ref}"${s}/>`;
    return `<c r="${ref}"${s}><v>${numberText(dateToSerial(value))}</v></c>`;
  }
  if (typeof value === 'number') {
    // A non-finite number (NaN, ±Infinity) has no OOXML representation; keep the cell and its style
    // but emit no value rather than a bare "NaN"/"Infinity" token — the same graceful degradation an
    // Invalid Date gets, so one bad value never corrupts the sheet or takes down the whole export.
    if (!Number.isFinite(value)) return `<c r="${ref}"${s}/>`;
    return `<c r="${ref}"${s}><v>${numberText(value)}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === 'string') {
    // With shared strings on, the cell holds only the pool index (`t="s"`); otherwise the text
    // lives inline in the cell. Both decode to the same string on read.
    if (sharedStrings !== null) {
      return `<c r="${ref}"${s} t="s"><v>${sharedStrings.intern(value)}</v></c>`;
    }
    return `<c r="${ref}"${s} t="inlineStr"><is>${textElement(value)}</is></c>`;
  }
  if (isRichTextValue(value)) {
    // With shared strings on, rich text is pooled as a rich `<si>` (the cell holds only its index);
    // otherwise the runs live inline. Both decode back to the same runs on read.
    if (sharedStrings !== null) {
      return `<c r="${ref}"${s} t="s"><v>${sharedStrings.intern(value)}</v></c>`;
    }
    return `<c r="${ref}"${s} t="inlineStr"><is>${richTextRunsXml(value.richText)}</is></c>`;
  }
  if (isHyperlinkValue(value)) {
    // The cell holds only the visible label; the link itself rides in the sheet's <hyperlinks>.
    // The label is either a plain string or rich text, serialised the same way a cell value of
    // that kind would be.
    const label =
      typeof value.text === 'string'
        ? textElement(value.text)
        : richTextRunsXml(value.text.richText);
    return `<c r="${ref}"${s} t="inlineStr"><is>${label}</is></c>`;
  }
  if (isErrorValue(value)) {
    // An error literal serialises under t="e" with its code as the value. The codes are a closed
    // set of canonical spellings (see ERROR_CODES) with no XML-special characters, so no escaping.
    return `<c r="${ref}"${s} t="e"><v>${value.error}</v></c>`;
  }
  // A null value only reaches here for a formatted-but-empty cell (the row loop keeps it for its
  // style); emit the styled cell with no <v>, exactly how Excel stores a formatted blank.
  if (value === null) return `<c r="${ref}"${s}/>`;
  throw new Error(`writing a ${detectValueType(value)} cell value is not implemented yet`);
}

// Whether a cell carries any style facet of its own — the reason to serialise it even when empty.
// A note is not a style: it lives in the comments part, not the cell's <c> element, so it does not
// count here. Row/column-inherited formatting is likewise excluded; only the cell's own facets do.
function hasOwnStyle(cell: Cell): boolean {
  return (
    cell.fill !== undefined ||
    cell.numFmt !== undefined ||
    cell.font !== undefined ||
    cell.border !== undefined ||
    cell.alignment !== undefined ||
    cell.protection !== undefined ||
    cell.quotePrefix === true ||
    cell.namedStyleId !== undefined
  );
}

function formulaCellXml(
  ref: string,
  s: string,
  formula: string,
  result: FormulaResult | undefined,
): string {
  return formulaBodyXml(ref, s, `<f>${escapeText(mangleFormula(formula))}</f>`, result);
}

// Wrap a prepared `<f>` element (a plain formula, or a shared master/slave `<f>`) with the cell
// element and its cached result, typing the cell by the result's kind exactly as a bare value of that
// kind would be.
function formulaBodyXml(
  ref: string,
  s: string,
  f: string,
  result: FormulaResult | undefined,
): string {
  // A non-finite cached result (a `1/0` that reached the model as Infinity/NaN) has no OOXML
  // representation; keep the formula but cache no value rather than emit a bare "NaN" — the same
  // graceful degradation a bare non-finite cell and an Invalid Date result get.
  if (result === undefined || (typeof result === 'number' && !Number.isFinite(result))) {
    return `<c r="${ref}"${s}>${f}</c>`;
  }
  if (typeof result === 'number') {
    return `<c r="${ref}"${s}>${f}<v>${numberText(result)}</v></c>`;
  }
  if (typeof result === 'boolean') {
    return `<c r="${ref}"${s} t="b">${f}<v>${result ? 1 : 0}</v></c>`;
  }
  if (typeof result === 'string') {
    return `<c r="${ref}"${s} t="str">${f}<v>${escapeText(result)}</v></c>`;
  }
  if (isErrorValue(result)) {
    // A formula that evaluated to an error caches its code under t="e", exactly as a bare error
    // cell does — the reader's decodeResult mirrors decodeValue for this case.
    return `<c r="${ref}"${s} t="e">${f}<v>${result.error}</v></c>`;
  }
  if (result instanceof Date) {
    // A date-valued result caches its serial exactly as a bare date cell stores its value; the
    // cell's date number format (applied when its style is composed) is what makes both read back as
    // a Date. An Invalid Date has no serial, so cache no result rather than emit NaN.
    if (Number.isNaN(result.getTime())) return `<c r="${ref}"${s}>${f}</c>`;
    return `<c r="${ref}"${s}>${f}<v>${numberText(dateToSerial(result))}</v></c>`;
  }
  // Every FormulaResult kind is handled above; this guards a value that reached here past the model.
  throw new Error('writing a non-primitive formula result is not implemented yet');
}

// A cell's role in an OOXML shared-formula group. A master carries the source formula plus the `ref`
// range the group spans; a clone (no `ref`) references the master's formula by the shared index `si`.
export interface SharedFormulaRole {
  readonly si: number;
  readonly ref?: string;
}

// Plan a sheet's shared-formula groups: every clone cell (a {@link SharedFormulaValue}) names its
// master by address, so group the clones by master, assign each group a sheet-unique `si`, and record
// the `ref` range (master through the furthest clone) on the master. Excel requires the master to sit
// at the top-left of that range, so a clone above or left of its master — or a master with no formula
// (an orphan) — is rejected here, named, rather than emitted as a package Excel repairs on open.
function planSharedFormulas(sheet: Worksheet): Map<string, SharedFormulaRole> {
  const groups = new Map<string, Cell[]>();
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      if (isSharedFormulaValue(cell.value)) {
        const clones = groups.get(cell.value.sharedFormula);
        if (clones !== undefined) clones.push(cell);
        else groups.set(cell.value.sharedFormula, [cell]);
      }
    }
  }

  const roles = new Map<string, SharedFormulaRole>();
  let si = 0;
  for (const [masterAddress, clones] of groups) {
    const master = sheet.getCell(masterAddress);
    if (!isFormulaValue(master.value)) {
      const offender = clones[0] as Cell;
      throw new Error(
        `shared-formula clone ${offender.address} names master ${masterAddress}, which holds no formula`,
      );
    }
    let maxCol = master.col;
    let maxRow = master.row;
    for (const clone of clones) {
      if (clone.col < master.col || clone.row < master.row) {
        throw new Error(
          `shared-formula master ${masterAddress} must sit above and/or left of clone ${clone.address}`,
        );
      }
      if (clone.col > maxCol) maxCol = clone.col;
      if (clone.row > maxRow) maxRow = clone.row;
    }
    roles.set(masterAddress, {
      si,
      ref: `${encodeAddress(master.col, master.row)}:${encodeAddress(maxCol, maxRow)}`,
    });
    for (const clone of clones) roles.set(clone.address, {si});
    si += 1;
  }
  return roles;
}
