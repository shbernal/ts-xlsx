// The buffered `.xlsx` writer: a Workbook model in, an OPC zip package out.
//
// It serialises the part of the model that exists today — worksheets; cells holding a
// number, string, boolean, or formula; column/row formatting; page margins and
// header/footer; merged ranges; and worksheet tables — into a valid package (content
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

import {strToU8, zipSync} from 'fflate';

import type {WorkbookImage} from '../../core/image.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {collectNotes, commentsXml, vmlDrawingXml} from './comments.ts';
import {collectHyperlinks, planHyperlinks} from './hyperlinks.ts';
import {drawingRelsXml, drawingXml} from './images.ts';
import {
  type BackgroundPlan,
  type CommentPlan,
  type DrawingPlan,
  type PivotPlan,
  type PlannedImage,
  type PlannedTable,
  type PreservedReferencePlan,
  type PrinterSettingsPlan,
  planMedia,
  planPreservedParts,
  SheetRelIds,
} from './package-plan.ts';
import {pivotCacheDefinitionXml, pivotCacheRecordsXml, pivotTableXml} from './pivot.ts';
import {REL, relsPartXml} from './relationships.ts';
import {SharedStringTable} from './shared-strings.ts';
import {THEME1_XML} from './static-parts.ts';
import {StyleRegistry} from './styles.ts';
import {
  appPropsXml,
  contentTypesXml,
  corePropsXml,
  rootRelsXml,
  workbookRelsXml,
  workbookXml,
} from './workbook-xml.ts';
import {
  type FlushedSheet,
  type SheetReferences,
  tableXml,
  worksheetRelsXml,
  worksheetXml,
} from './worksheet-xml.ts';

export {
  buildColumnDefaults,
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

  /**
   * The style registry to intern into, in place of a freshly-seeded one. The streaming writer
   * serialises each committed row eagerly (freeing its cells), so those rows' style ids must be
   * assigned by the very same registry that later emits `xl/styles.xml` — otherwise the ids in the
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
 * @throws {Error} if the workbook has no worksheets (a zero-sheet package is corrupt),
 *   or holds a value the writer cannot yet represent.
 */
export function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array {
  return zipSync(buildPackageParts(workbook, options), {level: 6});
}

/**
 * A style registry seeded from a workbook's read-in style layers (differential styles, named cell
 * styles, custom indexed palette), ready to intern authored styles after them. Both the buffered
 * pass and the streaming writer build their registry through here so a cell's style id means the
 * same thing whichever writer emits it.
 */
export function createStyleRegistry(workbook: Workbook): StyleRegistry {
  const styles = new StyleRegistry();
  // Seed the differential-style table with the fragments read from a source file so conditional
  // formatting's dxfId references stay valid; styles authored on rules append after them.
  styles.seedDifferentialStyles(workbook.differentialStyles);
  // Seed the named cell-style layer (cellStyleXfs/cellStyles) so each style's facets re-intern into the
  // rebuilt sub-tables and a cell's xfId link stays valid; without any, the default Normal alone emits.
  styles.seedNamedStyles(workbook.namedStyles);
  // Seed the custom indexed-color palette so it re-emits verbatim and an `indexed="…"` colour keeps
  // its intended RGB; a workbook that never overrode the palette seeds nothing and writes no <colors>.
  styles.seedIndexedColors(workbook.indexedColors);
  return styles;
}

/**
 * Assemble a workbook into the map of OPC package parts (part name → bytes) that make up an `.xlsx`,
 * short of zipping them. This is the whole serialisation — content types, relationships, workbook,
 * per-sheet XML, styles, theme, media, tables, and props — factored out of {@link writeXlsx} so the
 * streaming writer can drive the identical parts through a streamed zip container rather than
 * `zipSync`. Neither writer duplicates a byte of serialisation.
 *
 * @throws {Error} if the workbook has no worksheets, or holds a value the writer cannot represent.
 */
export function buildPackageParts(
  workbook: Workbook,
  options: WriteOptions = {},
): Record<string, Uint8Array> {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) {
    throw new Error(
      'cannot write a workbook with no worksheets — a zero-sheet package is corrupt to Excel',
    );
  }

  // With the option on, plain string cell values are pooled into a shared-strings table interned
  // during the sheet pass (like the style registry); a null table keeps every string inline.
  const sharedStrings = options.useSharedStrings ? new SharedStringTable() : null;

  // Anchored images share workbook-wide media: every image a sheet references becomes one media part,
  // addressed by a global number. Resolved before the sheet loop so a drawing's embeds can target it.
  const media = planMedia(workbook, sheets);

  // Content the model does not interpret — a vector-shape drawing, a header/footer image, a pivot
  // table and its caches, a slicer — captured on read and re-emitted verbatim onto collision-proof
  // paths. Preserved parts are renumbered past the parts the writer generates of the same kind
  // (drawings, VML, media), so resolving them needs only those generated counts; each sheet's
  // preserved references take their sheet-local rel ids in canonical position in the loop below.
  const generatedDrawingCount = sheets.filter((sheet) => sheet.images.length > 0).length;
  const preserved = planPreservedParts(workbook, generatedDrawingCount, media.parts.length);

  // Plan every sheet's parts in a single pass, drawing each sheet-local relationship id from that
  // sheet's allocator in canonical order: tables, drawing, comments (VML + comments part), printer
  // settings, external hyperlinks, background, preserved references, pivot tables. One running
  // allocator per sheet is what keeps the ids gapless and collision-free — no step re-derives its
  // offset by summing the ones before it, so none can drift into another's id. Part numbers (tables,
  // drawings, pivots) are global across the workbook and counted here in the same pass.
  let tableNumber = 0;
  let drawingNumber = 0;
  let pivotNumber = 0;
  const perSheet = sheets.map((sheet, i) => {
    const rels = new SheetRelIds();

    const tables: PlannedTable[] = sheet.tables.map((table) => ({
      table,
      number: ++tableNumber,
      relId: rels.next(),
    }));

    let drawing: DrawingPlan | null = null;
    if (sheet.images.length > 0) {
      const images: PlannedImage[] = sheet.images.map((image, j) => {
        const registered = workbook.getImage(image.imageId) as WorkbookImage;
        return {
          anchor: image.anchor,
          // The embed id is local to the drawing part's own rels, not the sheet's, so it is numbered
          // per image from rId1 rather than drawn from the sheet allocator.
          embedId: `rId${j + 1}`,
          mediaNumber: media.numberById.get(image.imageId) as number,
          extension: registered.extension,
        };
      });
      drawing = {number: ++drawingNumber, relId: rels.next(), images};
    }

    const notes = collectNotes(sheet);
    const comments: CommentPlan | null =
      notes.length === 0
        ? null
        : {number: i + 1, notes, vmlRelId: rels.next(), commentsRelId: rels.next()};

    const printerData = sheet.pageSetup.printerSettings;
    const printerSettings: PrinterSettingsPlan | null =
      printerData === undefined ? null : {number: i + 1, data: printerData, relId: rels.next()};

    const hyperlinks = planHyperlinks(collectHyperlinks(sheet), rels);

    let background: BackgroundPlan | null = null;
    if (sheet.backgroundImageId !== undefined) {
      const registered = workbook.getImage(sheet.backgroundImageId);
      if (registered === undefined) {
        throw new Error(
          `sheet "${sheet.name}" sets background image id ${sheet.backgroundImageId}, which is not registered on the workbook`,
        );
      }
      background = {
        relId: rels.next(),
        mediaNumber: media.numberById.get(sheet.backgroundImageId) as number,
        extension: registered.extension,
      };
    }

    const preservedRefs: PreservedReferencePlan[] = (preserved.perSheet[i] ?? []).map(
      (reference) => ({...reference, relId: rels.next()}),
    );

    const pivots: PivotPlan[] = sheet.pivotTables.map((table) => {
      const number = ++pivotNumber;
      // Each pivot is numbered globally (its parts and its `cacheId` must be workbook-unique); the
      // workbook relationship reaching its cache is assigned once the modeled workbook rels are known.
      return {number, cacheId: String(number), table, sheetRelId: rels.next(), workbookRelId: ''};
    });

    return {
      tables,
      drawing,
      comments,
      printerSettings,
      hyperlinks,
      background,
      preservedRefs,
      pivots,
    };
  });

  const sheetTables = perSheet.map((plan) => plan.tables);
  const sheetDrawings = perSheet.map((plan) => plan.drawing);
  const sheetComments = perSheet.map((plan) => plan.comments);
  const sheetPrinterSettings = perSheet.map((plan) => plan.printerSettings);
  const sheetHyperlinks = perSheet.map((plan) => plan.hyperlinks);
  const sheetBackgrounds = perSheet.map((plan) => plan.background);
  const sheetPreservedRefs = perSheet.map((plan) => plan.preservedRefs);
  const sheetPivots = perSheet.map((plan) => plan.pivots);
  const allTables = sheetTables.flat();
  const allPivots = sheetPivots.flat();

  // Serialise the worksheets first: interning each cell/row fill into the style table is a
  // side effect of that pass, so styles.xml can only be generated once every sheet is done. The
  // streaming writer supplies its own registry (already seeded, and already carrying its eagerly
  // flushed rows' styles); the buffered path seeds a fresh one here.
  const styles = options.styles ?? createStyleRegistry(workbook);
  const sheetXml = sheets.map((sheet, i) => {
    const refs = sheetPreservedRefs[i] ?? [];
    // A preserved `<drawing>` and a modeled one are mutually exclusive (a drawing is only preserved
    // when the sheet modeled no image from it), so the `<drawing>` slot takes whichever exists.
    const preservedDrawingRelId = refs.find((ref) => ref.element === 'drawing')?.relId ?? null;
    const legacyDrawingHFRelId =
      refs.find((ref) => ref.element === 'legacyDrawingHF')?.relId ?? null;
    // A slicer is wired into the sheet body by an `<x14:slicerList>` extension that names the same
    // relationship id the sheet's slicer rel carries — re-emitting it reactivates the widget rather
    // than leaving the preserved slicer part orphaned.
    const slicerRelIds = refs
      .filter((ref) => ref.relType.endsWith('/slicer'))
      .map((ref) => ref.relId);
    const references: SheetReferences = {
      drawingRelId: sheetDrawings[i]?.relId ?? preservedDrawingRelId,
      legacyDrawingRelId: sheetComments[i]?.vmlRelId ?? null,
      printerSettingsRelId: sheetPrinterSettings[i]?.relId ?? null,
      backgroundRelId: sheetBackgrounds[i]?.relId ?? null,
      legacyDrawingHFRelId,
      slicerRelIds,
    };
    return worksheetXml(
      sheet,
      sheetTables[i] ?? [],
      styles,
      references,
      sheetHyperlinks[i] ?? [],
      sharedStrings,
      options.flushed?.get(sheet),
    );
  });

  // The pool is filled only once every sheet is serialised. Emit the part (and its rel + content
  // type) solely when the option is on and at least one string was interned, so a workbook with no
  // string cells never fabricates an empty table.
  const hasSharedStrings = sharedStrings !== null && !sharedStrings.isEmpty;

  const commentNumbers = sheetComments
    .filter((c): c is CommentPlan => c !== null)
    .map((c) => c.number);
  const drawingNumbers = sheetDrawings
    .filter((d): d is DrawingPlan => d !== null)
    .map((d) => d.number);
  const printerSettingsNumbers = sheetPrinterSettings
    .filter((p): p is PrinterSettingsPlan => p !== null)
    .map((p) => p.number);

  // A preserved workbook reference's relationship id follows the modeled workbook rels — the sheets,
  // styles, theme, and (when emitted) shared strings — so adding one never renumbers an id already
  // used. The workbook body and its rels part are wired from the same assignment, so a pivot cache's
  // `<pivotCaches>` registration and its relationship agree on the id.
  const workbookRelBase = sheets.length + 2 + (hasSharedStrings ? 1 : 0);
  const preservedWorkbookRels = preserved.workbook.map((ref, i) => ({
    ...ref,
    relId: `rId${workbookRelBase + 1 + i}`,
  }));
  // A generated pivot cache's workbook relationship follows the preserved ones; the assignment
  // mutates the shared plan so the `<pivotCaches>` body and the rels part read the same id.
  const pivotWorkbookRelBase = workbookRelBase + preserved.workbook.length;
  allPivots.forEach((pivot, i) => {
    pivot.workbookRelId = `rId${pivotWorkbookRelBase + 1 + i}`;
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
      ),
    ),
    '_rels/.rels': strToU8(rootRelsXml()),
    'docProps/core.xml': strToU8(corePropsXml(workbook.properties)),
    'docProps/app.xml': strToU8(appPropsXml()),
    'xl/workbook.xml': strToU8(workbookXml(workbook, preservedWorkbookRels, allPivots)),
    'xl/_rels/workbook.xml.rels': strToU8(
      workbookRelsXml(sheets.length, hasSharedStrings, preservedWorkbookRels, allPivots),
    ),
    'xl/styles.xml': strToU8(styles.toXml()),
    'xl/theme/theme1.xml': strToU8(THEME1_XML),
  };
  if (hasSharedStrings) {
    files['xl/sharedStrings.xml'] = strToU8((sharedStrings as SharedStringTable).toXml());
  }
  for (const part of media.parts) {
    files[`xl/media/image${part.number}.${part.extension}`] = part.data;
  }
  sheets.forEach((_sheet, i) => {
    const tables = sheetTables[i] ?? [];
    const drawing = sheetDrawings[i] ?? null;
    const comments = sheetComments[i] ?? null;
    const printerSettings = sheetPrinterSettings[i] ?? null;
    const background = sheetBackgrounds[i] ?? null;
    const hyperlinks = sheetHyperlinks[i] ?? [];
    const preservedRefs = sheetPreservedRefs[i] ?? [];
    const pivots = sheetPivots[i] ?? [];
    const hasExternalHyperlink = hyperlinks.some((link) => link.relId !== undefined);
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml[i] as string);
    if (
      tables.length > 0 ||
      drawing !== null ||
      comments !== null ||
      printerSettings !== null ||
      background !== null ||
      hasExternalHyperlink ||
      preservedRefs.length > 0 ||
      pivots.length > 0
    ) {
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(
        worksheetRelsXml(
          tables,
          drawing,
          comments,
          printerSettings,
          background,
          hyperlinks,
          preservedRefs,
          pivots,
        ),
      );
    }
    if (printerSettings !== null) {
      files[`xl/printerSettings/printerSettings${printerSettings.number}.bin`] =
        printerSettings.data;
    }
    if (drawing !== null) {
      files[`xl/drawings/drawing${drawing.number}.xml`] = strToU8(drawingXml(drawing.images));
      const targets = drawing.images.map(
        (image) => `../media/image${image.mediaNumber}.${image.extension}`,
      );
      files[`xl/drawings/_rels/drawing${drawing.number}.xml.rels`] = strToU8(
        drawingRelsXml(targets),
      );
    }
    if (comments !== null) {
      files[`xl/comments${comments.number}.xml`] = strToU8(commentsXml(comments.notes));
      files[`xl/drawings/vmlDrawing${comments.number}.vml`] = strToU8(
        vmlDrawingXml(comments.notes),
      );
    }
  });
  for (const {table, number} of allTables) {
    files[`xl/tables/table${number}.xml`] = strToU8(tableXml(table, number));
  }
  // A pivot spans three parts wired in a chain: the pivot-table part (linked from its host sheet)
  // references the cache definition, which references the cache records. Each cache carries a rels
  // part naming the next link by `rId1` — the id the definition/table XML resolves against.
  for (const pivot of allPivots) {
    const {number, cacheId, table} = pivot;
    files[`xl/pivotTables/pivotTable${number}.xml`] = strToU8(
      pivotTableXml(table, `PivotTable${number}`, cacheId),
    );
    files[`xl/pivotTables/_rels/pivotTable${number}.xml.rels`] = strToU8(
      relsPartXml([
        {
          id: 'rId1',
          type: REL.pivotCacheDefinition,
          target: `../pivotCache/pivotCacheDefinition${number}.xml`,
        },
      ]),
    );
    files[`xl/pivotCache/pivotCacheDefinition${number}.xml`] = strToU8(
      pivotCacheDefinitionXml(table),
    );
    files[`xl/pivotCache/_rels/pivotCacheDefinition${number}.xml.rels`] = strToU8(
      relsPartXml([
        {id: 'rId1', type: REL.pivotCacheRecords, target: `pivotCacheRecords${number}.xml`},
      ]),
    );
    files[`xl/pivotCache/pivotCacheRecords${number}.xml`] = strToU8(pivotCacheRecordsXml(table));
  }
  // Emit the verbatim-preserved parts (and their rewired rels) last: their paths are collision-proof,
  // so ordering against the generated parts does not matter.
  for (const part of preserved.parts) {
    files[part.path] = part.bytes;
    if (part.relsPath !== null && part.relsXml !== null) {
      files[part.relsPath] = strToU8(part.relsXml);
    }
  }

  return files;
}
