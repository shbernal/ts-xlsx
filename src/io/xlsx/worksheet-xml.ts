// Worksheet serialisation: a Worksheet model into its `xl/worksheets/sheetN.xml` part and the sheet's
// rels part. Owns the row/cell renderer the streaming writer also drives, and orchestrates the whole
// sheet body — the print/page/view/protection blocks live in `sheet-properties.ts`, shared-formula
// planning in `shared-formulas.ts`, each imported here rather than duplicated. Table *parts*
// (`xl/tables/tableN.xml`) are `tables.ts`'s concern, alongside their reader — this module only wires
// the sheet's `<tableParts>` back-references to them.

import {decodeRange, encodeAddress, MAX_COLUMN} from '../../core/address.ts';
import type {Cell} from '../../core/cell.ts';
import {DEFAULT_DATE_NUMFMT, dateToSerial} from '../../core/date.ts';
import {mangleFormula} from '../../core/formula.ts';
import type {Fill} from '../../core/style.ts';
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
  RowProperties,
  Worksheet,
  WorksheetProperties,
} from '../../core/worksheet.ts';
import {escapeAttr, escapeText, numberText, textElement, XML_DECLARATION} from '../../xml/xml.ts';
import {relativePartPath} from '../opc/part-paths.ts';
import {relationship, relationshipsPart} from '../opc/rels.ts';
import {conditionalFormattingsExtXml, conditionalFormattingsXml} from './conditional-formatting.ts';
import {dataValidationsExtXml, dataValidationsXml} from './data-validation.ts';
import {type HyperlinkPlan, hyperlinksXml} from './hyperlinks.ts';
import {SLICER_LIST_EXT_URI} from './namespaces.ts';
import type {
  BackgroundPlan,
  CommentPlan,
  DrawingPlan,
  PivotPlan,
  PreservedReferencePlan,
  PrinterSettingsPlan,
  TablePlan,
  ThreadedCommentPlan,
} from './package-plan.ts';
import {NS, REL} from './relationships.ts';
import {richTextRunsXml} from './rich-text.ts';
import {planSharedFormulas, type SharedFormulaRole} from './shared-formulas.ts';
import type {SharedStringTable} from './shared-strings.ts';
import {
  autoFilterXml,
  headerFooterXml,
  pageBreaksXml,
  pageMarginsXml,
  pageSetupXml,
  printOptionsXml,
  sheetProtectionXml,
  sheetPrXml,
  sheetViewsXml,
} from './sheet-properties.ts';
import type {CellStyle, StyleRegistry} from './styles.ts';
import {x14Ext} from './x14-ext.ts';

/**
 * The used-cell extent of a sheet — the top-left/bottom-right grid bounds that fold into the
 * `<dimension>`. Rows carrying only formatting (a row height, an outline level) do not extend the
 * used range, matching how Excel records `<dimension>`, so {@link add} ignores them. A fresh extent
 * holds the `Infinity`/`-Infinity` sentinels; {@link isEmpty} reports that no used cell has been seen.
 */
export class Extent {
  top = Infinity;
  left = Infinity;
  bottom = -Infinity;
  right = -Infinity;

  // Seed from a prior extent (the rows a streaming writer already flushed and evicted) so the buffered
  // pass folds its live rows onto the same bounds; unseeded, it starts empty.
  constructor(seed?: Extent) {
    if (seed) {
      this.top = seed.top;
      this.left = seed.left;
      this.bottom = seed.bottom;
      this.right = seed.right;
    }
  }

  /** Whether no used cell has been folded in yet — the sheet's dimension is then the lone cell `A1`. */
  get isEmpty(): boolean {
    return this.bottom === -Infinity;
  }

  /** Fold a rendered row's used-column span into the extent. `minCol` is `Infinity` when the row
   * carried no cells (only formatting), which extends nothing. */
  add(row: number, minCol: number, maxCol: number): void {
    if (minCol === Infinity) return;
    if (row < this.top) this.top = row;
    if (row > this.bottom) this.bottom = row;
    if (minCol < this.left) this.left = minCol;
    if (maxCol > this.right) this.right = maxCol;
  }
}

/**
 * A worksheet's eagerly-serialised rows: each row's `<row>` XML tagged with its number (so it merges
 * into ascending order with the sheet's remaining live rows, whatever order it was committed in), plus
 * the used-cell {@link Extent} they span. The buffered pass folds that extent into the sheet's dimension.
 */
export interface FlushedSheet {
  readonly rows: ReadonlyArray<{readonly number: number; readonly xml: string}>;
  readonly extent: Extent;
  /**
   * The deepest row outline level among the flushed rows. Carried across the eviction because
   * `<sheetFormatPr outlineLevelRow>` is derived from every row on the sheet, and a flushed row's
   * properties are gone from the model by the time the header is serialised.
   */
  readonly maxRowOutlineLevel: number;
}

// The sheet-local relationship ids that wire a worksheet's tail elements to their parts, gathered into
// one struct so the caller hands them over as a named unit rather than a run of positional `string |
// null` arguments a mis-ordered call could silently transpose. Each id is `null` when the sheet carries
// no part of that kind. `slicerRelIds` is plural because a sheet may carry several preserved slicers.
export interface SheetReferences {
  readonly drawingRelId: string | null;
  readonly legacyDrawingRelId: string | null;
  readonly printerSettingsRelId: string | null;
  readonly backgroundRelId: string | null;
  readonly legacyDrawingHFRelId: string | null;
  readonly slicerRelIds: readonly string[];
}

export function worksheetXml(
  sheet: Worksheet,
  tables: readonly TablePlan[],
  styles: StyleRegistry,
  references: SheetReferences,
  hyperlinks: readonly HyperlinkPlan[],
  sharedStrings: SharedStringTable | null,
  active: boolean,
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
  // so the row loop can stamp it even onto a summary row that carries no properties of its own. The
  // same pass yields the sheet's deepest row outline level for `<sheetFormatPr>`.
  const rowOutline = scanRowOutline(sheet);
  const collapsedSummaries = rowOutline.collapsedSummaries;

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
  const extent = new Extent(flushed?.extent);

  for (const entry of sheet.rows()) {
    const {xml, minCol, maxCol} = renderRow(entry, context);
    if (xml === '') continue;
    liveRows.push({number: entry.number, xml});
    extent.add(entry.number, minCol, maxCol);
  }

  const dimensionRef = extent.isEmpty
    ? 'A1'
    : `${encodeAddress(extent.left, extent.top)}:${encodeAddress(extent.right, extent.bottom)}`;
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
    sheetViewsXml(sheet.view, active) +
    sheetFormatPr(sheet.properties, {
      col: maxColumnOutlineLevel(sheet),
      // A streamed sheet's flushed rows are gone from the model; their deepest level rides along on
      // the flush record so the header still reports the whole sheet's outline.
      row: Math.max(rowOutline.maxLevel, flushed?.maxRowOutlineLevel ?? 0),
    }) +
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
    pageSetupXml(sheet.pageSetup, references.printerSettingsRelId) +
    headerFooterXml(sheet.headerFooter) +
    // CT_Worksheet order: <rowBreaks> follows <headerFooter>, <colBreaks> follows <rowBreaks>, and
    // both precede the drawing block.
    pageBreaksXml(sheet.rowBreaks, 'rowBreaks') +
    pageBreaksXml(sheet.columnBreaks, 'colBreaks') +
    // Schema order near the tail: <drawing> (the images), then <legacyDrawing> (the VML holding the
    // note boxes), then <legacyDrawingHF> (a preserved header/footer image's VML), then <picture>
    // (the sheet background), then <tableParts>.
    refElement('drawing', references.drawingRelId) +
    refElement('legacyDrawing', references.legacyDrawingRelId) +
    refElement('legacyDrawingHF', references.legacyDrawingHFRelId) +
    refElement('picture', references.backgroundRelId) +
    tablePartsXml(tables) +
    // `<extLst>` is the final child of CT_Worksheet and a worksheet may carry at most one. Both the
    // x14 conditional-formatting extensions (data-bar gradient/negative-fill/axis) and the extended
    // (x14) data validations ride inside it as sibling `<ext>` blocks — so they are gathered here into
    // a single `<extLst>` rather than each emitting its own.
    worksheetExtLstXml(sheet, references.slicerRelIds) +
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
      const style = ctx.styles.styleId(
        composeCellStyle(cell, rowFill, ctx.columnDefaults.get(cell.col)),
      );
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

// Compose a cell's full style by resolving each facet cell-over-row-over-column, so a cell that
// overrides one facet still carries the row's fill and the column's other facets rather than silently
// dropping them — the per-facet precedence Excel applies. The row contributes only a fill today;
// quote-prefix and the named-style link are cell-only, with no row/column default to inherit.
function composeCellStyle(
  cell: Cell,
  rowFill: Fill | undefined,
  colDef: ColumnProperties | undefined,
): CellStyle {
  return {
    fill: cell.fill ?? rowFill ?? colDef?.fill,
    // A bare Date carries no format of its own, so it renders as a raw serial and reads back as a
    // number unless we apply a date format. An explicit cell/column format wins.
    numFmt: cell.numFmt ?? colDef?.numFmt ?? dateDefaultNumFmt(cell.value),
    font: cell.font ?? colDef?.font,
    border: cell.border ?? colDef?.border,
    alignment: cell.alignment ?? colDef?.alignment,
    protection: cell.protection ?? colDef?.protection,
    quotePrefix: cell.quotePrefix,
    // Preserved so a round-trip keeps the cell tied to its named style rather than flattening it into
    // a purely-direct format.
    xfId: cell.namedStyleId,
  };
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
  return x14Ext(SLICER_LIST_EXT_URI, `<x14:slicerList>${slicers}</x14:slicerList>`);
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
          `merged range ${merge} overlaps table "${table.name}" (${table.range}) — Excel forbids a merge inside a table`,
        );
      }
    }
  }
}

function mergeCellsXml(merges: readonly string[]): string {
  if (merges.length === 0) return '';
  const cells = merges
    .map((range) => `<mergeCell ref="${escapeAttr(decodeRange(range).dimensions)}"/>`)
    .join('');
  return `<mergeCells count="${merges.length}">${cells}</mergeCells>`;
}

// A tail reference element (`<drawing r:id="…"/>` and its `<legacyDrawing>`/`<legacyDrawingHF>`/
// `<picture>` siblings) wiring the sheet to a part by relationship id, or '' when the sheet carries no
// part of that kind — each such id is null then.
function refElement(tag: string, relId: string | null): string {
  return relId === null ? '' : `<${tag} r:id="${relId}"/>`;
}

function tablePartsXml(tables: readonly TablePlan[]): string {
  if (tables.length === 0) return '';
  const parts = tables.map(({relId}) => `<tablePart r:id="${relId}"/>`).join('');
  return `<tableParts count="${tables.length}">${parts}</tableParts>`;
}

export function worksheetRelsXml(
  tables: readonly TablePlan[],
  drawing: DrawingPlan | null,
  comments: CommentPlan | null,
  threadedComments: ThreadedCommentPlan | null,
  printerSettings: PrinterSettingsPlan | null,
  background: BackgroundPlan | null,
  hyperlinks: readonly HyperlinkPlan[],
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
    // A threaded-comment part, like a pivot table, is reached by relationship alone — no worksheet element
    // names it, so this relationship is the only thing that makes Excel look for the conversation.
    ...(threadedComments === null
      ? []
      : [
          relationship(
            threadedComments.relId,
            REL.threadedComment,
            `../threadedComments/threadedComment${threadedComments.number}.xml`,
          ),
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
    // TargetMode="External". Internal links have no relId and contribute nothing here.
    ...hyperlinks
      .filter((link) => link.relId !== undefined && link.target !== undefined)
      .map((link) =>
        relationship(link.relId as string, REL.hyperlink, escapeAttr(link.target as string), {
          external: true,
        }),
      ),
  ];
  return relationshipsPart(rels);
}

// Excel's standard row height in points, emitted as the `defaultRowHeight` when the sheet does not
// override it so a reader sees the same baseline Excel would write.
const DEFAULT_ROW_HEIGHT = 15;

// `<sheetFormatPr>` carries the sheet's grid defaults and, when the sheet groups anything, the depth
// of its deepest outline. A consumer sizes the outline bars from those depths — the strips that sit
// above the column headers and left of the row headers — so a grouped sheet that omits them lays its
// grid out with no room reserved for a bar it then has to draw. Both are omitted at zero, as Excel
// does, so an ungrouped sheet stays byte-clean.
function sheetFormatPr(
  properties: WorksheetProperties,
  outlineLevel: {readonly col: number; readonly row: number},
): string {
  const rowHeight = properties.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
  let attrs = ` defaultRowHeight="${numberText(rowHeight)}"`;
  if (properties.defaultColWidth !== undefined) {
    attrs += ` defaultColWidth="${numberText(properties.defaultColWidth)}"`;
  }
  // A non-standard default row height is only honoured by Excel when customHeight is set.
  if (properties.defaultRowHeight !== undefined) attrs += ' customHeight="1"';
  if (outlineLevel.col > 0) attrs += ` outlineLevelCol="${outlineLevel.col}"`;
  if (outlineLevel.row > 0) attrs += ` outlineLevelRow="${outlineLevel.row}"`;
  return `<sheetFormatPr${attrs}/>`;
}

// The deepest column outline level the sheet declares — the `outlineLevelCol` its `<sheetFormatPr>`
// reports. A column past XFD contributes nothing: {@link colsXml} drops it as out-of-range, so its
// group would have no `<col>` to sit on.
function maxColumnOutlineLevel(sheet: Worksheet): number {
  let max = 0;
  for (const {index, properties} of sheet.columns()) {
    if (index > MAX_COLUMN) continue;
    max = Math.max(max, properties.outlineLevel ?? 0);
  }
  return max;
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

// The two row-outline facts the serialiser needs, from one walk over the rows: which summary rows
// terminate a fully-collapsed group, and how deep the sheet's grouping goes. They share a pass
// because the pass is the expensive part — the streaming writer must not be made to traverse rows
// twice just to fill in a header attribute.
interface RowOutline {
  readonly collapsedSummaries: Set<number>;
  readonly maxLevel: number;
}

// A collapsed outline group is two coordinated facts: its detail rows carry outlineLevel and are
// hidden, AND the summary row that terminates the group carries `collapsed`. Authors typically set
// only outlineLevel + hidden on the detail rows, so the summary flag is derived here rather than
// demanded of the caller: a row is a collapsed summary iff its adjacent detail run — the contiguous
// higher-outline-level rows on the summary side — is non-empty and every row in it is hidden.
// Placement follows the sheet's summaryBelow flag (Excel's default is summary below the detail); the
// walk stops at the first row of level <= the summary's own, so a gap or a boundary ends the group.
function scanRowOutline(sheet: Worksheet): RowOutline {
  const level = new Map<number, number>();
  const hidden = new Map<number, boolean>();
  let maxLevel = 0;
  for (const {number, properties} of sheet.rows()) {
    const rowLevel = properties?.outlineLevel ?? 0;
    level.set(number, rowLevel);
    hidden.set(number, properties?.hidden ?? false);
    if (rowLevel > maxLevel) maxLevel = rowLevel;
  }
  const levelOf = (row: number): number => level.get(row) ?? 0;
  const step = sheet.outline.summaryBelow === false ? 1 : -1;
  const collapsedSummaries = new Set<number>();
  for (const [summary, summaryLevel] of level) {
    let detail = summary + step;
    let sawDetail = false;
    let allHidden = true;
    while (levelOf(detail) > summaryLevel) {
      sawDetail = true;
      if (!hidden.get(detail)) allHidden = false;
      detail += step;
    }
    if (sawDetail && allHidden) collapsedSummaries.add(summary);
  }
  return {collapsedSummaries, maxLevel};
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

  const formula = cellFormulaXml(ref, s, value, shared);
  if (formula !== undefined) return formula;

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

// Serialise a formula cell — a shared-formula master or clone, a What-If data table, or a plain
// formula — into its `<c>` element, or return undefined when the value is not a formula so `cellXml`
// falls through to its value dispatch.
function cellFormulaXml(
  ref: string,
  s: string,
  value: Cell['value'],
  shared: SharedFormulaRole | undefined,
): string | undefined {
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
    return formulaBodyXml(
      ref,
      s,
      `<f>${escapeText(mangleFormula(value.formula))}</f>`,
      value.result,
    );
  }
  return undefined;
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
