// The buffered `.xlsx` writer: a Workbook model in, an OPC zip package out.
//
// It serialises the part of the model that exists today (worksheets; cells holding a
// number, string, boolean, or formula; column/row formatting; page margins and
// header/footer; merged ranges; and worksheet tables) into a valid package (content
// types, relationships, workbook, per-sheet XML, table parts, the default theme and
// stylesheet, and core/app properties). Styles, images, and the richer value kinds land
// as the model grows; until then the writer refuses a value it cannot represent
// faithfully rather than emitting a lossy or corrupt package.
//
// This module is the orchestrator: it plans the package graph (via `package-plan.ts`) and
// stitches the serialised parts (from `workbook-xml.ts` and `worksheet-xml.ts`) into the
// part map. The row/cell renderer and the sheet's public render types live in
// `worksheet-xml.ts` and are re-exported here so the streaming writer's import surface is
// unchanged.

import {strToU8, zip, zipSync} from 'fflate';

import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {relativePartPath, relsPathFor, THEME_PART_PATH} from '../opc/part-paths.ts';
import {relsPartXml} from '../opc/rels.ts';
import {FIXED_ENTRY_MTIME} from '../opc/zip-mtime.ts';
import {collectComments, commentsXml, vmlDrawingXml} from './comments.ts';
import {collectHyperlinks, type HyperlinkPlan, planHyperlinks} from './hyperlinks.ts';
import {drawingRelsXml, drawingXml} from './images.ts';
import {
  type BackgroundPlan,
  type CommentPlan,
  type DrawingPlan,
  type ImagePlan,
  type MediaPlan,
  type PivotPlan,
  type PreservedPartPlan,
  type PreservedPlan,
  type PreservedReferencePlan,
  type PreservedWorkbookReferencePlan,
  type PrinterSettingsPlan,
  planMedia,
  planPreservedParts,
  SheetRelIds,
  type TablePlan,
  type ThreadedCommentPlan,
} from './package-plan.ts';
import {
  APP_PROPS_PART,
  commentsPart,
  CORE_PROPS_PART,
  drawingPart,
  mediaPart,
  PERSONS_PART,
  pivotCacheDefinitionPart,
  pivotCacheRecordsPart,
  pivotTablePart,
  printerSettingsPart,
  SHARED_STRINGS_PART,
  STYLES_PART,
  tablePart,
  threadedCommentsPart,
  vmlDrawingPart,
  WORKBOOK_PART,
  worksheetPart,
} from './part-names.ts';
import {pivotCacheDefinitionXml, pivotCacheRecordsXml, pivotTableXml} from './pivot.ts';
import {REL} from './relationships.ts';
import {SharedStringTable} from './shared-strings.ts';
import {StyleRegistry} from './styles.ts';
import {tableXml} from './tables.ts';
import {applyThemeOverrides, DEFAULT_THEME_XML} from './theme-xml.ts';
import {personsXml, threadedCommentsXml} from './threaded-comments.ts';
import {
  appPropsXml,
  contentTypesXml,
  corePropsXml,
  FIXED_WORKBOOK_REL_COUNT,
  rootRelsXml,
  workbookRelsXml,
  workbookXml,
} from './workbook-xml.ts';
import {
  type FlushedSheet,
  type SheetReferences,
  worksheetRelsXml,
  worksheetXml,
} from './worksheet-xml.ts';

export {
  buildColumnDefaults,
  Extent,
  type FlushedSheet,
  type RowRenderContext,
  renderRow,
} from './worksheet-xml.ts';

/** Options controlling how {@link writeXlsx} serialises a workbook. */
export interface WriteOptions {
  /**
   * Pool plain string cell values into a shared-strings table (`xl/sharedStrings.xml`) that cells
   * reference by index, rather than storing each string inline in its cell. Deduplicates repeated
   * text and matches Excel's own storage; off by default, which keeps strings inline and omits the
   * part. Rich-text values stay inline regardless, so their run formatting is unaffected.
   */
  readonly useSharedStrings?: boolean;
}

/**
 * {@link WriteOptions} plus the streaming writer's internal wiring, so a buffered caller's options
 * object can never carry fields meant only for {@link WorkbookStreamWriter}'s own use. Not exported
 * from the public barrel: {@link WorkbookStreamWriter} is the only caller that fills these fields,
 * and it reaches them through {@link buildPackageParts}.
 */
export interface InternalWriteOptions extends WriteOptions {
  /**
   * The style registry to intern into, in place of a freshly-seeded one. The streaming writer
   * serialises each committed row eagerly (freeing its cells), so those rows' style ids must be
   * assigned by the very same registry that later emits `xl/styles.xml`; otherwise the ids in the
   * pre-rendered rows would not match the styles part. When omitted the buffered path seeds its own,
   * so its output is unchanged.
   */
  readonly styles?: StyleRegistry;

  /**
   * Per-sheet rows already serialised and evicted from the model by the streaming writer, keyed by
   * the model worksheet. Their XML is emitted ahead of the sheet's remaining live rows and their
   * extent folds into `<dimension>`. Absent for the buffered path, which holds every row live.
   */
  readonly flushed?: ReadonlyMap<Worksheet, FlushedSheet>;
}

/**
 * Serialise a workbook into an `.xlsx` package.
 *
 * The bytes are a pure function of the workbook: an unchanged model written twice produces two
 * identical archives, because entry timestamps are pinned to a fixed date rather than taken from the
 * clock. A committed `.xlsx` therefore only changes when something about it changed.
 *
 * @throws {AuthoringError} if the workbook has no worksheets (a zero-sheet package is corrupt),
 *   or holds a value the writer cannot yet represent.
 */
export function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array {
  return zipSync(buildPackageParts(workbook, options), {level: 6, mtime: FIXED_ENTRY_MTIME});
}

/**
 * Serialise a workbook into an `.xlsx` package, deflating off the calling thread.
 *
 * Produces the same package {@link writeXlsx} does, byte for byte and entry timestamps included, and
 * exists for one reason: DEFLATE dominates the cost of writing a large workbook, and {@link writeXlsx}
 * spends all of it on the caller's thread. Here `fflate` deflates each part in a worker, so the event
 * loop keeps turning (stalls drop from the whole write to tens of milliseconds) and parts compress in
 * parallel, which on a multi-sheet workbook also finishes sooner. On a single-sheet workbook there is
 * only one part to deflate, so expect responsiveness rather than speed.
 *
 * Building the parts still happens on the calling thread; only compression moves. That is why there
 * is no `readXlsxAsync` mirroring this: reading is dominated by XML parsing and model building, which
 * no worker can take, and the reader's zip-bomb ceiling is enforced by counting output between
 * synchronous input slices. See ADR-0024.
 *
 * @throws {AuthoringError}, as a rejection, under the same conditions as {@link writeXlsx};
 *   the part-building it shares happens before any worker is involved. A failure raised by the zip
 *   layer itself (including an environment that cannot spawn a worker) propagates unwrapped, exactly
 *   as it does from {@link writeXlsx}.
 */
export async function writeXlsxAsync(
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  const parts = buildPackageParts(workbook, options);
  return await new Promise<Uint8Array>((resolve, reject) => {
    zip(parts, {level: 6, mtime: FIXED_ENTRY_MTIME}, (error, data) => {
      if (error) reject(error);
      else resolve(data);
    });
  });
}

/**
 * A style registry seeded from a workbook's read-in style layers (differential styles, named cell
 * styles, custom indexed palette), ready to intern authored styles after them. Both the buffered
 * pass and the streaming writer build their registry through here so a cell's style id means the
 * same thing whichever writer emits it.
 */
export function createStyleRegistry(workbook: Workbook): StyleRegistry {
  // Font id 0 is the workbook's own default face, resolved from what it declared, what was authored,
  // and its theme's body typeface. Never an assumed Calibri, which would re-face every empty cell
  // and change the metric every character-unit column width is expressed in.
  const styles = new StyleRegistry({
    defaultFont: workbook.defaultFont,
    ...(workbook.declaredDefaultFont === undefined
      ? {}
      : {declaredDefaultFont: workbook.declaredDefaultFont}),
  });
  // Seed the differential-style table with the fragments read from a source file so conditional
  // formatting's dxfId references stay valid; styles authored on rules append after them.
  styles.seedDifferentialStyles(workbook.differentialStyles);
  // Seed the named cell-style layer (cellStyleXfs/cellStyles) so each style's facets re-intern into the
  // rebuilt sub-tables and a cell's xfId link stays valid; without any, the default Normal alone emits.
  styles.seedNamedStyles(workbook.namedStyles);
  // Seed the custom indexed-color palette so it re-emits verbatim and an `indexed="…"` colour keeps
  // its intended RGB; a workbook that never overrode the palette seeds nothing and writes no <colors>.
  styles.seedIndexedColors(workbook.indexedColors);
  // Seed the author's "Recent Colors" swatches so they re-emit unchanged rather than being reset.
  styles.seedMruColors(workbook.mruColors);
  // Seed the custom table-style definitions so a table's `styleName` still names a real definition and
  // each element's dxfId still indexes the differential-style table seeded above at its original index.
  styles.seedTableStyles(workbook.tableStyles);
  // Authored styles append after the preserved ones, and intern their elements' formatting after the
  // seeded dxfs: the ordering that keeps every preserved dxfId pointing where it did.
  for (const style of workbook.customTableStyles) styles.addTableStyle(style);
  return styles;
}

// One worksheet's planned package parts and the sheet-local relationship ids wiring them, produced in
// the single planning pass. Held as a struct per sheet rather than eight index-aligned arrays, so a
// downstream step reads one sheet's plan as a unit and cannot transpose two sheets by mis-indexing.
interface SheetPlan {
  readonly tables: TablePlan[];
  readonly drawing: DrawingPlan | null;
  readonly comments: CommentPlan | null;
  readonly threadedComments: ThreadedCommentPlan | null;
  readonly printerSettings: PrinterSettingsPlan | null;
  readonly hyperlinks: HyperlinkPlan[];
  readonly background: BackgroundPlan | null;
  readonly preservedRefs: PreservedReferencePlan[];
  readonly pivots: PivotPlan[];
}

// The workbook-global part counters the per-sheet planning advances: tables, drawings and pivot
// tables are numbered across the whole package, not per sheet. Passed as one mutable object so the
// shared state is part of {@link planSheet}'s signature rather than three variables it closes over.
interface PartNumbering {
  table: number;
  drawing: number;
  pivot: number;
}

/**
 * Plan one sheet's parts, drawing every sheet-local relationship id from that sheet's own allocator
 * in the one canonical order the package wires them: tables, drawing, comments (VML + comments part),
 * threaded comments, printer settings, external hyperlinks, background, preserved references, pivot
 * tables.
 *
 * That order is the correctness story. One running allocator per sheet is what keeps the ids gapless
 * and collision-free, because no step re-derives its offset by summing the ones before it, so none
 * can drift into another's id. A step moved above another here silently renumbers a package, and
 * nothing but the corpus would say so.
 */
function planSheet(context: {
  readonly sheet: Worksheet;
  readonly index: number;
  readonly media: MediaPlan;
  readonly preserved: PreservedPlan;
  readonly numbering: PartNumbering;
}): SheetPlan {
  const {sheet, index, media, preserved, numbering} = context;
  const rels = new SheetRelIds();

  const tables: TablePlan[] = sheet.tables.map((table) => ({
    table,
    number: ++numbering.table,
    relId: rels.next(),
  }));

  let drawing: DrawingPlan | null = null;
  if (sheet.images.length > 0) {
    const images: ImagePlan[] = sheet.images.map((image, j) => {
      const {number, image: registered} = media.resolve(image.imageId);
      return {
        anchor: image.anchor,
        // The embed id is local to the drawing part's own rels, not the sheet's, so it is numbered
        // per image from rId1 rather than drawn from the sheet allocator.
        embedId: `rId${j + 1}`,
        mediaNumber: number,
        extension: registered.extension,
      };
    });
    drawing = {number: ++numbering.drawing, relId: rels.next(), images};
  }

  // A conversation and the legacy fallback `<comment>` that binds its cell to it are two halves of one
  // representation, so both are derived from this single list and neither can be emitted without the
  // other. Verified against desktop Excel: a `tc=` fallback whose thread part is absent shows as neither
  // a thread nor a note (the text disappears rather than degrading) and a thread part whose fallback is
  // absent is ignored, leaving the cell blank. A thread with no messages is not one of them: it has
  // nothing to say, and no head id for its replies or its fallback to hang off.
  const threads = sheet.commentThreads.filter((thread) => thread.comments.length > 0);
  const sheetComments = collectComments(sheet, threads);
  const comments: CommentPlan | null =
    sheetComments.length === 0
      ? null
      : {
          number: index + 1,
          comments: sheetComments,
          vmlRelId: rels.next(),
          commentsRelId: rels.next(),
        };
  const threadedComments: ThreadedCommentPlan | null =
    threads.length === 0 ? null : {number: index + 1, threads, relId: rels.next()};

  const printerData = sheet.pageSetup.printerSettings;
  const printerSettings: PrinterSettingsPlan | null =
    printerData === undefined ? null : {number: index + 1, data: printerData, relId: rels.next()};

  const hyperlinks = planHyperlinks(collectHyperlinks(sheet), rels);

  let background: BackgroundPlan | null = null;
  if (sheet.backgroundImageId !== undefined) {
    const {number, image} = media.resolve(sheet.backgroundImageId);
    background = {relId: rels.next(), mediaNumber: number, extension: image.extension};
  }

  const preservedRefs: PreservedReferencePlan[] = (preserved.perSheet[index] ?? []).map(
    (reference) => ({...reference, relId: rels.next()}),
  );

  const pivots: PivotPlan[] = sheet.pivotTables.map((table) => {
    const number = ++numbering.pivot;
    // Each pivot is numbered globally (its parts and its `cacheId` must be workbook-unique); the
    // workbook relationship reaching its cache is assigned once the modeled workbook rels are known.
    return {number, cacheId: String(number), table, sheetRelId: rels.next(), workbookRelId: ''};
  });

  return {
    tables,
    drawing,
    comments,
    threadedComments,
    printerSettings,
    hyperlinks,
    background,
    preservedRefs,
    pivots,
  };
}

// The part numbers of whichever sheets carry a part of one kind. Four call sites spelled this as a
// map/filter/map triple with a type predicate whose only job was to restate the plan type it had just
// selected; one picker says the same thing without the predicate.
function numbersOf(
  plans: readonly SheetPlan[],
  pick: (plan: SheetPlan) => {readonly number: number} | null,
): number[] {
  const numbers: number[] = [];
  for (const plan of plans) {
    const part = pick(plan);
    if (part !== null) numbers.push(part.number);
  }
  return numbers;
}

// Resolve one sheet's tail reference ids (the `<drawing>`/`<legacyDrawing>`/`<legacyDrawingHF>`/
// `<picture>` slots and the slicer list) from its plan. A preserved `<drawing>` and a modeled one are
// mutually exclusive, so the drawing slot takes whichever exists; a comment's VML rides the legacy-
// drawing slot; and each preserved slicer surfaces its rel id so the `<x14:slicerList>` can reactivate
// the widget rather than orphan its part.
function resolveSheetReferences(plan: SheetPlan): SheetReferences {
  const refs = plan.preservedRefs;
  const preservedDrawingRelId = refs.find((ref) => ref.element === 'drawing')?.relId ?? null;
  const legacyDrawingHFRelId = refs.find((ref) => ref.element === 'legacyDrawingHF')?.relId ?? null;
  const slicerRelIds = refs
    .filter((ref) => ref.relType.endsWith('/slicer'))
    .map((ref) => ref.relId);
  return {
    drawingRelId: plan.drawing?.relId ?? preservedDrawingRelId,
    legacyDrawingRelId: plan.comments?.vmlRelId ?? null,
    printerSettingsRelId: plan.printerSettings?.relId ?? null,
    backgroundRelId: plan.background?.relId ?? null,
    legacyDrawingHFRelId,
    slicerRelIds,
  };
}

/**
 * Lay the workbook-level relationship ids out in one place, and return the two the caller needs to
 * thread onward.
 *
 * The order is the whole of it, and it is the subtlest arithmetic in the writer. The modeled rels
 * come first (one per sheet, then the fixed styles/theme pair, then shared strings when emitted),
 * because they are the ones an existing package already numbered: anything laid after them can be
 * added without renumbering an id already in use. The threaded-comment person registry follows, then
 * the preserved workbook references, then the generated pivot caches.
 *
 * A pivot's id is written back onto the shared {@link PivotPlan} rather than returned, because two
 * separate parts have to agree on it: the workbook body's `<pivotCaches>` registration and the rels
 * part. Wiring both from one assignment is what makes disagreeing impossible; returning it would put
 * the burden back on two call sites to use the same value.
 */
function assignWorkbookRelIds(context: {
  readonly sheetCount: number;
  readonly hasSharedStrings: boolean;
  readonly hasPersons: boolean;
  readonly preservedWorkbook: readonly PreservedWorkbookReferencePlan[];
  readonly pivots: readonly PivotPlan[];
}): {
  readonly personsRelId: string | null;
  readonly preservedWorkbookRels: readonly (PreservedWorkbookReferencePlan & {relId: string})[];
} {
  const {sheetCount, hasSharedStrings, hasPersons, preservedWorkbook, pivots} = context;
  const modeledCount = sheetCount + FIXED_WORKBOOK_REL_COUNT + (hasSharedStrings ? 1 : 0);
  const personsRelId = hasPersons ? `rId${modeledCount + 1}` : null;
  const preservedBase = modeledCount + (personsRelId === null ? 0 : 1);
  const preservedWorkbookRels = preservedWorkbook.map((ref, i) => ({
    ...ref,
    relId: `rId${preservedBase + 1 + i}`,
  }));
  const pivotBase = preservedBase + preservedWorkbook.length;
  pivots.forEach((pivot, i) => {
    pivot.workbookRelId = `rId${pivotBase + 1 + i}`;
  });
  return {personsRelId, preservedWorkbookRels};
}

// Everything about a package that is resolved before any of its bytes exist: the media every sheet
// shares, the parts carried over verbatim, and each sheet's own parts with their relationship ids
// already allocated. Named because it is the boundary between the two halves of the writer: nothing
// above it serialises, and nothing below it decides what the package contains.
interface PackagePlan {
  readonly media: MediaPlan;
  readonly preserved: PreservedPlan;
  readonly perSheet: readonly SheetPlan[];
  readonly allTables: readonly TablePlan[];
  readonly allPivots: readonly PivotPlan[];
}

// Resolve the whole package graph: the media the sheets share, the verbatim-preserved parts numbered
// past the generated ones, then every sheet's parts in a single pass through {@link planSheet}.
function planPackage(workbook: Workbook, sheets: readonly Worksheet[]): PackagePlan {
  // Anchored images share workbook-wide media: every image a sheet references becomes one media part,
  // addressed by a global number. Resolved before the sheet loop so a drawing's embeds can target it.
  const media = planMedia(workbook, sheets);

  // Content the model does not interpret (a vector-shape drawing, a header/footer image, a pivot
  // table and its caches, a slicer) captured on read and re-emitted verbatim onto collision-proof
  // paths. Preserved parts are renumbered past the parts the writer generates of the same kind
  // (drawings, VML, media), so resolving them needs only those generated counts; each sheet's
  // preserved references take their sheet-local rel ids in canonical position below.
  const generatedDrawingCount = sheets.filter((sheet) => sheet.images.length > 0).length;
  const preserved = planPreservedParts(workbook, generatedDrawingCount, media.parts.length);

  // The part numbers that are global across the workbook (tables, drawings, pivots) run through one
  // shared counter, which is why this is a `map` over a mutable object rather than a pure one: the
  // shared state is in the signature instead of being three `let`s a callback happens to close over.
  const numbering: PartNumbering = {table: 0, drawing: 0, pivot: 0};
  const perSheet = sheets.map((sheet, index) =>
    planSheet({sheet, index, media, preserved, numbering}),
  );

  return {
    media,
    preserved,
    perSheet,
    allTables: perSheet.flatMap((plan) => plan.tables),
    allPivots: perSheet.flatMap((plan) => plan.pivots),
  };
}

/**
 * Serialise every worksheet, in sheet order.
 *
 * This runs before `xl/styles.xml` is generated, and must: interning a cell's or row's format into
 * the style table is a side effect of this pass, so the stylesheet is only complete once every sheet
 * has been through it. Emitting the styles part first would silently drop the styles of whatever had
 * not been serialised yet.
 */
function serialiseSheets(
  workbook: Workbook,
  sheets: readonly Worksheet[],
  plan: PackagePlan,
  styles: StyleRegistry,
  sharedStrings: SharedStringTable | null,
  flushed: InternalWriteOptions['flushed'],
): string[] {
  return sheets.map((sheet, i) => {
    const sheetPlan = plan.perSheet[i] as SheetPlan;
    return worksheetXml(
      sheet,
      sheetPlan.tables,
      styles,
      resolveSheetReferences(sheetPlan),
      sheetPlan.hyperlinks,
      sharedStrings,
      // Exactly one sheet is marked selected; the model resolves which, so no package can ship with
      // none selected (no view initialised on open) or with several (an accidental group selection,
      // where an edit to one sheet lands on all of them).
      i === workbook.activeTabIndex,
      flushed?.get(sheet),
    );
  });
}

/**
 * Build the part map itself: every part path paired with its bytes.
 *
 * The workbook-level relationship ids are laid out here rather than in the plan because two of their
 * inputs are known only once the sheets are serialised: whether any string was interned into the
 * shared-strings pool, and whether any sheet carried a conversation for the person registry to serve.
 */
function emitPackageParts(context: {
  readonly workbook: Workbook;
  readonly sheets: readonly Worksheet[];
  readonly plan: PackagePlan;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly sheetXml: readonly string[];
}): Record<string, Uint8Array> {
  const {workbook, sheets, plan, styles, sharedStrings, sheetXml} = context;
  const {media, preserved, perSheet, allTables, allPivots} = plan;

  // The pool is filled only once every sheet is serialised. Emit the part (and its rel + content
  // type) solely when the option is on and at least one string was interned, so a workbook with no
  // string cells never fabricates an empty table.
  const hasSharedStrings = sharedStrings !== null && !sharedStrings.isEmpty;

  const commentNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.comments);
  const drawingNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.drawing);
  const printerSettingsNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.printerSettings);
  const threadedCommentNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.threadedComments);

  // The identity registry is emitted only beside the thread parts that point into it. With no conversation
  // in the package nothing can reference a `<person>`, so the part would be a workbook-level relationship
  // to dead weight, and it is the messages, not the registry, that make an identity worth carrying.
  const persons = threadedCommentNumbers.length === 0 ? [] : workbook.persons;

  const {personsRelId, preservedWorkbookRels} = assignWorkbookRelIds({
    sheetCount: sheets.length,
    hasSharedStrings,
    hasPersons: persons.length > 0,
    preservedWorkbook: preserved.workbook,
    pivots: allPivots,
  });

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      contentTypesXml(
        sheets.length,
        allTables,
        commentNumbers,
        drawingNumbers,
        printerSettingsNumbers,
        media.extensions,
        hasSharedStrings,
        preserved.parts,
        allPivots,
        preservedWorkbookRels,
        threadedCommentNumbers,
        persons.length > 0,
      ),
    ),
    '_rels/.rels': strToU8(rootRelsXml(preserved.root)),
    [CORE_PROPS_PART]: strToU8(corePropsXml(workbook.properties)),
    [APP_PROPS_PART]: strToU8(appPropsXml(workbook.properties)),
    [WORKBOOK_PART]: strToU8(workbookXml(workbook, preservedWorkbookRels, allPivots)),
    [relsPathFor(WORKBOOK_PART)]: strToU8(
      workbookRelsXml(
        sheets.length,
        hasSharedStrings,
        personsRelId,
        preservedWorkbookRels,
        allPivots,
      ),
    ),
    [STYLES_PART]: strToU8(styles.toXml()),
  };
  // A theme read from a source package is emitted through the preserved-part path, closure and all,
  // with any authored overrides already spliced into its entry part by the planner. A workbook without
  // one gets its authored theme, or the library's default, which the stylesheet's `theme="1"` default
  // font still needs something to resolve against.
  if (!preserved.themeEmitted) {
    const overrides = workbook.themeOverrides;
    files[THEME_PART_PATH] = strToU8(
      overrides === undefined
        ? DEFAULT_THEME_XML
        : applyThemeOverrides(DEFAULT_THEME_XML, overrides),
    );
  }
  if (hasSharedStrings) {
    files[SHARED_STRINGS_PART] = strToU8(sharedStrings.toXml());
  }
  // Singular and unnumbered, unlike the per-sheet thread parts: one registry serves the whole workbook.
  if (persons.length > 0) files[PERSONS_PART] = strToU8(personsXml(persons));
  for (const part of media.parts) {
    files[mediaPart(part.number, part.extension)] = part.data;
  }
  emitSheetParts(files, perSheet, sheetXml);
  for (const {table, number} of allTables) {
    files[tablePart(number)] = strToU8(tableXml(table, number));
  }
  emitPivotParts(files, allPivots);
  emitPreservedParts(files, preserved.parts);

  return files;
}

/**
 * Assemble a workbook into the map of OPC package parts (part name → bytes) that make up an `.xlsx`,
 * short of zipping them. This is the whole serialisation (content types, relationships, workbook,
 * per-sheet XML, styles, theme, media, tables, and props) factored out of {@link writeXlsx} so the
 * streaming writer can drive the identical parts through a streamed zip container rather than
 * `zipSync`. Neither writer duplicates a byte of serialisation.
 *
 * Three steps, and their order is load-bearing. The package is planned in full before anything is
 * serialised, so a drawing's embed and a pivot's cache can target a part number that already exists.
 * The sheets are serialised before the parts are emitted, because interning a style is a side effect
 * of that pass and `xl/styles.xml` is one of the parts emitted in the third step.
 *
 * @throws {AuthoringError} if the workbook has no worksheets, or holds a value the writer cannot represent.
 */
export function buildPackageParts(
  workbook: Workbook,
  options: InternalWriteOptions = {},
): Record<string, Uint8Array> {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) {
    throw new AuthoringError(
      'cannot write a workbook with no worksheets: a zero-sheet package is corrupt to Excel',
    );
  }

  // With the option on, plain string cell values are pooled into a shared-strings table interned
  // during the sheet pass (like the style registry); a null table keeps every string inline.
  const sharedStrings = options.useSharedStrings ? new SharedStringTable() : null;
  // The streaming writer supplies its own registry (already seeded, and already carrying its eagerly
  // flushed rows' styles); the buffered path seeds a fresh one here.
  const styles = options.styles ?? createStyleRegistry(workbook);

  const plan = planPackage(workbook, sheets);
  const sheetXml = serialiseSheets(workbook, sheets, plan, styles, sharedStrings, options.flushed);
  return emitPackageParts({workbook, sheets, plan, styles, sharedStrings, sheetXml});
}

// The map of OPC part paths to their serialised bytes, accumulated by {@link buildPackageParts} and
// its per-phase `emit*` helpers.
type PackageFiles = Record<string, Uint8Array>;

// Emit each sheet's own parts: the sheet XML, its rels part (only when the sheet references something),
// and the drawing/comment/printer-settings parts those relationships point at. `sheetXml[i]` is the
// already-serialised body for `perSheet[i]`, indexed in lockstep.
function emitSheetParts(
  files: PackageFiles,
  perSheet: readonly SheetPlan[],
  sheetXml: readonly string[],
): void {
  perSheet.forEach((plan, i) => {
    const {
      tables,
      drawing,
      comments,
      threadedComments,
      printerSettings,
      background,
      hyperlinks,
      preservedRefs,
      pivots,
    } = plan;
    const hasExternalHyperlink = hyperlinks.some((link) => link.relId !== undefined);
    files[worksheetPart(i + 1)] = strToU8(sheetXml[i] as string);
    if (
      tables.length > 0 ||
      drawing !== null ||
      comments !== null ||
      threadedComments !== null ||
      printerSettings !== null ||
      background !== null ||
      hasExternalHyperlink ||
      preservedRefs.length > 0 ||
      pivots.length > 0
    ) {
      files[relsPathFor(worksheetPart(i + 1))] = strToU8(
        worksheetRelsXml(
          tables,
          drawing,
          comments,
          threadedComments,
          printerSettings,
          background,
          hyperlinks,
          preservedRefs,
          pivots,
        ),
      );
    }
    if (printerSettings !== null) {
      files[printerSettingsPart(printerSettings.number)] = printerSettings.data;
    }
    if (drawing !== null) {
      const drawingPath = drawingPart(drawing.number);
      files[drawingPath] = strToU8(drawingXml(drawing.images));
      const targets = drawing.images.map((image) =>
        relativePartPath(drawingPath, mediaPart(image.mediaNumber, image.extension)),
      );
      files[relsPathFor(drawingPath)] = strToU8(drawingRelsXml(targets));
    }
    if (comments !== null) {
      files[commentsPart(comments.number)] = strToU8(commentsXml(comments.comments));
      files[vmlDrawingPart(comments.number)] = strToU8(vmlDrawingXml(comments.comments));
    }
    if (threadedComments !== null) {
      files[threadedCommentsPart(threadedComments.number)] = strToU8(
        threadedCommentsXml(threadedComments.threads),
      );
    }
  });
}

// Emit every pivot table's three chained parts. A pivot spans a pivot-table part (linked from its host
// sheet) that references a cache definition, which references its cache records. Each cache carries a
// rels part naming the next link by `rId1`: the id the definition/table XML resolves against.
function emitPivotParts(files: PackageFiles, allPivots: readonly PivotPlan[]): void {
  for (const pivot of allPivots) {
    const {number, cacheId, table} = pivot;
    const tablePath = pivotTablePart(number);
    const definitionPath = pivotCacheDefinitionPart(number);
    files[tablePath] = strToU8(pivotTableXml(table, `PivotTable${number}`, cacheId));
    files[relsPathFor(tablePath)] = strToU8(
      relsPartXml([
        {
          id: 'rId1',
          type: REL.pivotCacheDefinition,
          target: relativePartPath(tablePath, definitionPath),
        },
      ]),
    );
    files[definitionPath] = strToU8(pivotCacheDefinitionXml(table));
    files[relsPathFor(definitionPath)] = strToU8(
      relsPartXml([
        {
          id: 'rId1',
          type: REL.pivotCacheRecords,
          target: relativePartPath(definitionPath, pivotCacheRecordsPart(number)),
        },
      ]),
    );
    files[pivotCacheRecordsPart(number)] = strToU8(pivotCacheRecordsXml(table));
  }
}

// Emit the verbatim-preserved parts (and their rewired rels) last: their paths are collision-proof, so
// ordering against the generated parts does not matter.
function emitPreservedParts(files: PackageFiles, parts: readonly PreservedPartPlan[]): void {
  for (const part of parts) {
    files[part.path] = part.bytes;
    if (part.relsPath !== null && part.relsXml !== null) {
      files[part.relsPath] = strToU8(part.relsXml);
    }
  }
}
