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
import {collectHyperlinks, type PlannedHyperlink, planHyperlinks} from './hyperlinks.ts';
import {drawingRelsXml, drawingXml} from './images.ts';
import {
  type BackgroundPlan,
  type CommentPlan,
  type DrawingPlan,
  type PivotPlan,
  type PlannedImage,
  type PlannedTable,
  type PrinterSettingsPlan,
  planMedia,
  planPreservedParts,
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
import {type FlushedSheet, tableXml, worksheetRelsXml, worksheetXml} from './worksheet-xml.ts';

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

  // Tables are numbered globally across the workbook (their part names and ids must be
  // unique), but relationship ids are local to each sheet's rels part.
  let tableNumber = 0;
  const sheetTables: PlannedTable[][] = sheets.map((sheet) =>
    sheet.tables.map((table, i) => ({table, number: ++tableNumber, relId: `rId${i + 1}`})),
  );
  const allTables = sheetTables.flat();

  // Anchored images share workbook-wide media: every image a sheet references becomes one media
  // part, addressed by a global number. A sheet's images then ride in a drawing part whose
  // sheet-local relationship id follows its table ids.
  const media = planMedia(workbook, sheets);
  let drawingNumber = 0;
  const sheetDrawings: (DrawingPlan | null)[] = sheets.map((sheet, i) => {
    if (sheet.images.length === 0) return null;
    const tableCount = (sheetTables[i] ?? []).length;
    const images: PlannedImage[] = sheet.images.map((image, j) => {
      const registered = workbook.getImage(image.imageId) as WorkbookImage;
      return {
        anchor: image.anchor,
        embedId: `rId${j + 1}`,
        mediaNumber: media.numberById.get(image.imageId) as number,
        extension: registered.extension,
      };
    });
    return {number: ++drawingNumber, relId: `rId${tableCount + 1}`, images};
  });

  // A sheet's notes ride in a comments part plus a legacy VML drawing. Their sheet-local
  // relationship ids follow the table and drawing ids, so a sheet with one table and a drawing
  // anchors its VML at rId3.
  const sheetComments: (CommentPlan | null)[] = sheets.map((sheet, i) => {
    const notes = collectNotes(sheet);
    if (notes.length === 0) return null;
    const base = (sheetTables[i] ?? []).length + (sheetDrawings[i] !== null ? 1 : 0);
    return {
      number: i + 1,
      notes,
      vmlRelId: `rId${base + 1}`,
      commentsRelId: `rId${base + 2}`,
    };
  });

  // A sheet's opaque printer-settings blob rides in its own binary part, linked from `<pageSetup>`
  // by a sheet-local relationship id that follows every other sheet-local id (tables, drawing,
  // comments) so adding one never renumbers the ids already threaded into the sheet XML.
  const sheetPrinterSettings: (PrinterSettingsPlan | null)[] = sheets.map((sheet, i) => {
    const data = sheet.pageSetup.printerSettings;
    if (data === undefined) return null;
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] !== null ? 1 : 0) +
      (sheetComments[i] !== null ? 2 : 0);
    return {number: i + 1, data, relId: `rId${base + 1}`};
  });

  // A sheet's hyperlinks add external relationships whose ids follow every other sheet-local id
  // (tables, drawing, comments, printer-settings), so introducing a link never renumbers an id
  // already threaded into the sheet XML. Internal ('#'-prefixed) links need no relationship at all.
  const sheetHyperlinks: PlannedHyperlink[][] = sheets.map((sheet, i) => {
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] !== null ? 1 : 0) +
      (sheetComments[i] !== null ? 2 : 0) +
      (sheetPrinterSettings[i] !== null ? 1 : 0);
    return planHyperlinks(collectHyperlinks(sheet), base);
  });

  // A sheet background is a workbook image referenced by a `<picture>` element through a sheet-local
  // image relationship. Its id follows every other sheet-local id (tables, drawing, comments,
  // printer-settings, external hyperlinks) so adding a background never renumbers an existing id.
  const sheetBackgrounds: (BackgroundPlan | null)[] = sheets.map((sheet, i) => {
    if (sheet.backgroundImageId === undefined) return null;
    const registered = workbook.getImage(sheet.backgroundImageId);
    if (registered === undefined) {
      throw new Error(
        `sheet "${sheet.name}" sets background image id ${sheet.backgroundImageId}, which is not registered on the workbook`,
      );
    }
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] !== null ? 1 : 0) +
      (sheetComments[i] !== null ? 2 : 0) +
      (sheetPrinterSettings[i] !== null ? 1 : 0) +
      (sheetHyperlinks[i] ?? []).filter((link) => link.relId !== undefined).length;
    return {
      relId: `rId${base + 1}`,
      mediaNumber: media.numberById.get(sheet.backgroundImageId) as number,
      extension: registered.extension,
    };
  });

  // Content the model does not interpret — a vector-shape drawing, a header/footer image, a pivot
  // table and its caches, a slicer — captured on read and re-emitted verbatim. Sheet references are
  // numbered after the backgrounds so their sheet-local rel ids sit at the tail of the sequence.
  const preserved = planPreservedParts(
    workbook,
    sheetTables,
    sheetDrawings,
    sheetComments,
    sheetPrinterSettings,
    sheetHyperlinks,
    sheetBackgrounds,
    media.parts.length,
  );

  // A pivot table hosted on a sheet adds one sheet-local relationship (to its pivot-table part),
  // numbered after every other sheet-local id — including the preserved refs at the tail — so
  // introducing a pivot never renumbers an id already threaded into the sheet or its rels. Each
  // pivot is numbered globally (its parts and its `cacheId` must be workbook-unique); the workbook
  // relationship reaching its cache is assigned once the modeled workbook rels are known (below).
  let pivotNumber = 0;
  const sheetPivots: PivotPlan[][] = sheets.map((sheet, i) => {
    if (sheet.pivotTables.length === 0) return [];
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] !== null ? 1 : 0) +
      (sheetComments[i] !== null ? 2 : 0) +
      (sheetPrinterSettings[i] !== null ? 1 : 0) +
      (sheetHyperlinks[i] ?? []).filter((link) => link.relId !== undefined).length +
      (sheetBackgrounds[i] !== null ? 1 : 0) +
      (preserved.perSheet[i]?.length ?? 0);
    return sheet.pivotTables.map((table, j) => {
      const number = ++pivotNumber;
      return {
        number,
        cacheId: String(number),
        table,
        sheetRelId: `rId${base + j + 1}`,
        workbookRelId: '',
      };
    });
  });
  const allPivots = sheetPivots.flat();

  // Serialise the worksheets first: interning each cell/row fill into the style table is a
  // side effect of that pass, so styles.xml can only be generated once every sheet is done. The
  // streaming writer supplies its own registry (already seeded, and already carrying its eagerly
  // flushed rows' styles); the buffered path seeds a fresh one here.
  const styles = options.styles ?? createStyleRegistry(workbook);
  const sheetXml = sheets.map((sheet, i) => {
    const refs = preserved.perSheet[i] ?? [];
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
    return worksheetXml(
      sheet,
      sheetTables[i] ?? [],
      styles,
      sheetDrawings[i]?.relId ?? preservedDrawingRelId,
      sheetComments[i]?.vmlRelId ?? null,
      sheetPrinterSettings[i]?.relId ?? null,
      sheetBackgrounds[i]?.relId ?? null,
      legacyDrawingHFRelId,
      slicerRelIds,
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
    const preservedRefs = preserved.perSheet[i] ?? [];
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
