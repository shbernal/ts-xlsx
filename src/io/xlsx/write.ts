// The buffered `.xlsx` writer: a Workbook model in, an OPC zip package out.
//
// It serialises the part of the model that exists today — worksheets; cells holding a
// number, string, boolean, or formula; column/row formatting; page margins and
// header/footer; merged ranges; and worksheet tables — into a valid package (content
// types, relationships, workbook, per-sheet XML, table parts, the default theme and
// stylesheet, and core/app properties). Styles, images, and the richer value kinds land
// as the model grows; until then the writer refuses a value it cannot represent
// faithfully rather than emitting a lossy or corrupt package.

import {zipSync, strToU8} from 'fflate';

import {decodeRange, encodeAddress, MAX_COLUMN} from '../../core/address.ts';
import type {AutoFilter, FilterColumn, FilterCriteria} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import {dateToSerial, DEFAULT_DATE_NUMFMT} from '../../core/date.ts';
import type {WorkbookImage} from '../../core/image.ts';
import {mangleFormula} from '../../core/formula.ts';
import type {PivotTable} from '../../core/pivot-table.ts';
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
import {type SheetProtection, SHEET_PROTECTION_FLAGS} from '../../core/protection.ts';
import type {Workbook, WorkbookProperties} from '../../core/workbook.ts';
import type {
  ColumnProperties,
  HeaderFooter,
  OutlineProperties,
  PageBreak,
  PageMargins,
  PageSetup,
  RowProperties,
  SheetView,
  Worksheet,
  WorksheetProperties,
} from '../../core/worksheet.ts';
import {collectNotes, commentsXml, type NoteCell, vmlDrawingXml} from './comments.ts';
import {conditionalFormattingsExtXml, conditionalFormattingsXml} from './conditional-formatting.ts';
import {dataValidationsExtXml, dataValidationsXml} from './data-validation.ts';
import {
  collectHyperlinks,
  hyperlinksXml,
  planHyperlinks,
  type PlannedHyperlink,
} from './hyperlinks.ts';
import {type DrawingImage, drawingRelsXml, drawingXml, imageContentType} from './images.ts';
import {pivotCacheDefinitionXml, pivotCacheRecordsXml, pivotTableXml} from './pivot.ts';
import {richTextRunsXml} from './rich-text.ts';
import {SharedStringTable} from './shared-strings.ts';
import {THEME1_XML} from './static-parts.ts';
import {colorAttrs, StyleRegistry} from './styles.ts';
import {escapeAttr, escapeText, textElement, XML_DECLARATION} from './xml.ts';

const NS = {
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  packageRels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  docRels: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
} as const;

const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  printerSettings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDefinition:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRecords:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
} as const;

const REL = {
  worksheet: `${NS.docRels}/worksheet`,
  styles: `${NS.docRels}/styles`,
  theme: `${NS.docRels}/theme`,
  officeDocument: `${NS.docRels}/officeDocument`,
  coreProps: `${NS.packageRels}/metadata/core-properties`,
  extProps: `${NS.docRels}/extended-properties`,
  table: `${NS.docRels}/table`,
  comments: `${NS.docRels}/comments`,
  vmlDrawing: `${NS.docRels}/vmlDrawing`,
  drawing: `${NS.docRels}/drawing`,
  printerSettings: `${NS.docRels}/printerSettings`,
  image: `${NS.docRels}/image`,
  hyperlink: `${NS.docRels}/hyperlink`,
  sharedStrings: `${NS.docRels}/sharedStrings`,
  pivotTable: `${NS.docRels}/pivotTable`,
  pivotCacheDefinition: `${NS.docRels}/pivotCacheDefinition`,
  pivotCacheRecords: `${NS.docRels}/pivotCacheRecords`,
} as const;

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
 * Serialise a workbook into an `.xlsx` package.
 *
 * @throws {Error} if the workbook has no worksheets (a zero-sheet package is corrupt),
 *   or holds a value the writer cannot yet represent.
 */
export function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array {
  return zipSync(buildPackageParts(workbook, options), {level: 6});
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
  options: WriteOptions = {}
): Record<string, Uint8Array> {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) {
    throw new Error('cannot write a workbook with no worksheets — a zero-sheet package is corrupt to Excel');
  }

  // With the option on, plain string cell values are pooled into a shared-strings table interned
  // during the sheet pass (like the style registry); a null table keeps every string inline.
  const sharedStrings = options.useSharedStrings ? new SharedStringTable() : null;

  // Tables are numbered globally across the workbook (their part names and ids must be
  // unique), but relationship ids are local to each sheet's rels part.
  let tableNumber = 0;
  const sheetTables: PlannedTable[][] = sheets.map(sheet =>
    sheet.tables.map((table, i) => ({table, number: ++tableNumber, relId: `rId${i + 1}`}))
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
        `sheet "${sheet.name}" sets background image id ${sheet.backgroundImageId}, which is not registered on the workbook`
      );
    }
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] !== null ? 1 : 0) +
      (sheetComments[i] !== null ? 2 : 0) +
      (sheetPrinterSettings[i] !== null ? 1 : 0) +
      (sheetHyperlinks[i] ?? []).filter(link => link.relId !== undefined).length;
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
    media.parts.length
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
      (sheetHyperlinks[i] ?? []).filter(link => link.relId !== undefined).length +
      (sheetBackgrounds[i] !== null ? 1 : 0) +
      (preserved.perSheet[i]?.length ?? 0);
    return sheet.pivotTables.map((table, j) => {
      const number = ++pivotNumber;
      return {number, cacheId: String(number), table, sheetRelId: `rId${base + j + 1}`, workbookRelId: ''};
    });
  });
  const allPivots = sheetPivots.flat();

  // Serialise the worksheets first: interning each cell/row fill into the style table is a
  // side effect of that pass, so styles.xml can only be generated once every sheet is done.
  const styles = new StyleRegistry();
  // Seed the differential-style table with the fragments read from a source file so conditional
  // formatting's dxfId references stay valid; styles authored on rules append after them.
  styles.seedDifferentialStyles(workbook.differentialStyles);
  // Seed the named cell-style layer (cellStyleXfs/cellStyles) so each style's facets re-intern into the
  // rebuilt sub-tables and a cell's xfId link stays valid; without any, the default Normal alone emits.
  styles.seedNamedStyles(workbook.namedStyles);
  const sheetXml = sheets.map((sheet, i) => {
    const refs = preserved.perSheet[i] ?? [];
    // A preserved `<drawing>` and a modeled one are mutually exclusive (a drawing is only preserved
    // when the sheet modeled no image from it), so the `<drawing>` slot takes whichever exists.
    const preservedDrawingRelId = refs.find(ref => ref.element === 'drawing')?.relId ?? null;
    const legacyDrawingHFRelId = refs.find(ref => ref.element === 'legacyDrawingHF')?.relId ?? null;
    // A slicer is wired into the sheet body by an `<x14:slicerList>` extension that names the same
    // relationship id the sheet's slicer rel carries — re-emitting it reactivates the widget rather
    // than leaving the preserved slicer part orphaned.
    const slicerRelIds = refs.filter(ref => ref.relType.endsWith('/slicer')).map(ref => ref.relId);
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
      sharedStrings
    );
  });

  // The pool is filled only once every sheet is serialised. Emit the part (and its rel + content
  // type) solely when the option is on and at least one string was interned, so a workbook with no
  // string cells never fabricates an empty table.
  const hasSharedStrings = sharedStrings !== null && !sharedStrings.isEmpty;

  const commentNumbers = sheetComments.filter((c): c is CommentPlan => c !== null).map(c => c.number);
  const drawingNumbers = sheetDrawings.filter((d): d is DrawingPlan => d !== null).map(d => d.number);
  const printerSettingsNumbers = sheetPrinterSettings
    .filter((p): p is PrinterSettingsPlan => p !== null)
    .map(p => p.number);

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
        allPivots
      )
    ),
    '_rels/.rels': strToU8(rootRelsXml()),
    'docProps/core.xml': strToU8(corePropsXml(workbook.properties)),
    'docProps/app.xml': strToU8(appPropsXml()),
    'xl/workbook.xml': strToU8(workbookXml(workbook, preservedWorkbookRels, allPivots)),
    'xl/_rels/workbook.xml.rels': strToU8(
      workbookRelsXml(sheets.length, hasSharedStrings, preservedWorkbookRels, allPivots)
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
        worksheetRelsXml(tables, drawing, comments, printerSettings, background, hyperlinks, preservedRefs, pivots)
      );
    }
    if (printerSettings !== null) {
      files[`xl/printerSettings/printerSettings${printerSettings.number}.bin`] = printerSettings.data;
    }
    if (drawing !== null) {
      files[`xl/drawings/drawing${drawing.number}.xml`] = strToU8(drawingXml(drawing.images));
      const targets = drawing.images.map(image => `../media/image${image.mediaNumber}.${image.extension}`);
      files[`xl/drawings/_rels/drawing${drawing.number}.xml.rels`] = strToU8(drawingRelsXml(targets));
    }
    if (comments !== null) {
      files[`xl/comments${comments.number}.xml`] = strToU8(commentsXml(comments.notes));
      files[`xl/drawings/vmlDrawing${comments.number}.vml`] = strToU8(vmlDrawingXml(comments.notes));
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
      pivotTableXml(table, `PivotTable${number}`, cacheId)
    );
    files[`xl/pivotTables/_rels/pivotTable${number}.xml.rels`] = strToU8(
      relsPartXml([
        {id: 'rId1', type: REL.pivotCacheDefinition, target: `../pivotCache/pivotCacheDefinition${number}.xml`},
      ])
    );
    files[`xl/pivotCache/pivotCacheDefinition${number}.xml`] = strToU8(pivotCacheDefinitionXml(table));
    files[`xl/pivotCache/_rels/pivotCacheDefinition${number}.xml.rels`] = strToU8(
      relsPartXml([{id: 'rId1', type: REL.pivotCacheRecords, target: `pivotCacheRecords${number}.xml`}])
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

// A pivot table planned for emission: its global part number, the workbook-unique `cacheId` its
// `<pivotCaches>` registration and `pivotTableDefinition` agree on, the sheet-local relationship
// linking its host sheet to the pivot-table part, and the workbook relationship reaching its cache
// definition (assigned once the modeled workbook rels are counted).
interface PivotPlan {
  readonly number: number;
  readonly cacheId: string;
  readonly table: PivotTable;
  readonly sheetRelId: string;
  workbookRelId: string;
}

// A sheet's notes paired with the part number and sheet-local relationship ids that link the sheet
// to its comments part (by type) and its VML drawing (by the `<legacyDrawing>` element).
interface CommentPlan {
  readonly number: number;
  readonly notes: readonly NoteCell[];
  readonly vmlRelId: string;
  readonly commentsRelId: string;
}

// A sheet's opaque printer-settings blob paired with the part number naming its `.bin` part and the
// sheet-local relationship id that links the sheet's `<pageSetup r:id>` to it.
interface PrinterSettingsPlan {
  readonly number: number;
  readonly data: Uint8Array;
  readonly relId: string;
}

// A table paired with the identifiers the package needs: a workbook-global part number
// and the sheet-local relationship id that links its worksheet to the table part.
interface PlannedTable {
  readonly table: Table;
  readonly number: number;
  readonly relId: string;
}

// A sheet background image resolved for serialisation: the sheet-local relationship id its
// `<picture>` element references, and the media part (global number + extension) that holds the bytes.
interface BackgroundPlan {
  readonly relId: string;
  readonly mediaNumber: number;
  readonly extension: string;
}

// A verbatim-preserved package part resolved for serialisation: the collision-proof path it is
// emitted at, its bytes and content type, and — when it references other parts — the rels part
// linking it to their new paths.
interface PreservedPartPlan {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly contentType: string;
  readonly relsPath: string | null;
  readonly relsXml: string | null;
}

// A preserved worksheet reference resolved for serialisation: the worksheet element that wires it
// (`null` for a pivot-table/slicer reference the sheet carries by relationship alone), the sheet-local
// relationship id, the relationship Type, and the new path of the entry part it targets.
interface PreservedReferencePlan {
  readonly element: 'drawing' | 'legacyDrawingHF' | null;
  readonly relId: string;
  readonly relType: string;
  readonly entryPath: string;
}

// A preserved workbook reference resolved for serialisation: its relationship Type, the new path of
// the entry part, and — for a pivot cache — the `cacheId` its `<pivotCaches>` registration carries.
// The workbook relationship id is assigned at emit time (it follows the modeled workbook rels, whose
// count depends on whether a shared-strings part is emitted), so it is not fixed here.
interface PreservedWorkbookReferencePlan {
  readonly relType: string;
  readonly entryPath: string;
  readonly pivotCacheId: string | undefined;
}

// The whole workbook's preserved content resolved for serialisation: per-sheet reference plans
// (parallel to the sheets) that wire the worksheet elements and rels, the workbook-level reference
// plans, and the flat, de-duplicated list of parts to emit. Kept together because the parts are
// numbered globally while the references are sheet- or workbook-local.
interface PreservedPlan {
  readonly perSheet: readonly (readonly PreservedReferencePlan[])[];
  readonly workbook: readonly PreservedWorkbookReferencePlan[];
  readonly parts: readonly PreservedPartPlan[];
}

// A sheet's drawing part: its workbook-global number, the sheet-local relationship id linking the
// sheet's `<drawing>` element to it, and the images it lays out.
interface DrawingPlan {
  readonly number: number;
  readonly relId: string;
  readonly images: readonly PlannedImage[];
}

// An anchored image resolved for serialisation: its anchor and drawing-local embed id (via
// DrawingImage) plus the media part number and extension its embed relationship targets.
interface PlannedImage extends DrawingImage {
  readonly mediaNumber: number;
  readonly extension: string;
}

// One picture written to `xl/media/`: its global part number, extension, and bytes.
interface MediaPart {
  readonly number: number;
  readonly extension: string;
  readonly data: Uint8Array;
}

// The workbook's media, resolved for writing: the parts to emit, a map from a workbook image id to
// its media part number (so a drawing embed can target it), and the distinct extensions in use (so
// content types can declare an image `<Default>` per extension).
interface MediaPlan {
  readonly parts: readonly MediaPart[];
  readonly numberById: ReadonlyMap<number, number>;
  readonly extensions: readonly string[];
}

// Gather the workbook images actually referenced by some sheet — either anchored in a drawing or set
// as a sheet background (an unreferenced image is not written) — number them in first-use order, and
// record the extensions in play. A sheet referencing an id with no registered image is a programming
// error the writer surfaces rather than emitting a dangling relationship.
function planMedia(workbook: Workbook, sheets: readonly Worksheet[]): MediaPlan {
  const usedIds: number[] = [];
  const seen = new Set<number>();
  const use = (id: number): void => {
    if (!seen.has(id)) {
      seen.add(id);
      usedIds.push(id);
    }
  };
  for (const sheet of sheets) {
    for (const image of sheet.images) use(image.imageId);
    if (sheet.backgroundImageId !== undefined) use(sheet.backgroundImageId);
  }
  const parts: MediaPart[] = [];
  const numberById = new Map<number, number>();
  const extensions = new Set<string>();
  usedIds.forEach((id, i) => {
    const image = workbook.getImage(id);
    if (image === undefined) {
      throw new Error(`a worksheet anchors image id ${id}, which is not registered on the workbook`);
    }
    const number = i + 1;
    parts.push({number, extension: image.extension, data: image.data});
    numberById.set(id, number);
    extensions.add(image.extension);
  });
  return {parts, numberById, extensions: [...extensions]};
}

// Resolve every sheet's verbatim-preserved worksheet references (a vector-shape drawing, a
// header/footer image) into the parts to emit and the sheet-local wiring to inject. Each reference's
// captured part closure is re-numbered onto collision-proof `preservedP{n}` paths — so preserved
// content never clobbers a generated drawing/VML/media part — with the closure's internal
// relationships rewritten to the new sibling paths. The worksheet element's relationship id follows
// every other sheet-local id (tables, drawing, comments, printer-settings, hyperlinks, background) so
// adding it never renumbers an id already threaded into the sheet XML.
function planPreservedParts(
  workbook: Workbook,
  sheetTables: readonly (readonly PlannedTable[])[],
  sheetDrawings: readonly (DrawingPlan | null)[],
  sheetComments: readonly (CommentPlan | null)[],
  sheetPrinterSettings: readonly (PrinterSettingsPlan | null)[],
  sheetHyperlinks: readonly (readonly PlannedHyperlink[])[],
  sheetBackgrounds: readonly (BackgroundPlan | null)[],
  generatedMediaCount: number
): PreservedPlan {
  const sheets = workbook.worksheets;
  // The writer generates drawings, VML, and media of its own, so a preserved part of one of those
  // kinds is re-numbered past the generated ones (a preserved drawing never clobbers an anchored
  // drawing, a preserved VML never clobbers a comment's VML). Comment VML is numbered by sheet index,
  // so `sheets.length` bounds it. Every other kind (pivot tables, caches, slicers, charts) the writer
  // never generates, so those keep their original path — see {@link preservedPartPath}.
  const numbering: PreservedNumbering = {
    drawing: sheetDrawings.filter((d): d is DrawingPlan => d !== null).length,
    vml: sheets.length,
    media: generatedMediaCount,
  };

  // One package-wide remap and one emitted-parts map: a part reached through more than one reference
  // (a pivot cache reached both from its pivot table and from the workbook) is numbered once and
  // emitted once, so overlapping closures collapse instead of duplicating parts.
  const remap = new Map<string, string>();
  const allReferences = [...sheets.flatMap(sheet => sheet.preservedReferences), ...workbook.preservedReferences];
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      if (!remap.has(part.path)) remap.set(part.path, preservedPartPath(part.path, numbering));
    }
  }
  const emitted = new Map<string, PreservedPartPlan>();
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      const newPath = remap.get(part.path) as string;
      if (emitted.has(newPath)) continue;
      const rels = part.rels.flatMap(rel => {
        const target = remap.get(rel.targetPath);
        return target === undefined ? [] : [{id: rel.id, type: rel.type, target: relativePartPath(newPath, target)}];
      });
      emitted.set(newPath, {
        path: newPath,
        bytes: part.bytes,
        contentType: part.contentType,
        relsPath: rels.length === 0 ? null : relsPathForPart(newPath),
        relsXml: rels.length === 0 ? null : preservedRelsXml(rels),
      });
    }
  }

  const perSheet = sheets.map((sheet, i): PreservedReferencePlan[] => {
    const references = sheet.preservedReferences;
    if (references.length === 0) return [];
    // A preserved reference's sheet-local relationship id follows every modeled sheet-local id
    // (tables, drawing, comments, printer-settings, external hyperlinks, background) so adding it never
    // renumbers an id already threaded into the sheet XML.
    const externalHyperlinks = (sheetHyperlinks[i] ?? []).filter(link => link.relId !== undefined).length;
    const base =
      (sheetTables[i] ?? []).length +
      (sheetDrawings[i] ? 1 : 0) +
      (sheetComments[i] ? 2 : 0) +
      (sheetPrinterSettings[i] ? 1 : 0) +
      externalHyperlinks +
      (sheetBackgrounds[i] ? 1 : 0);
    let next = base + 1;
    return references.map((reference): PreservedReferencePlan => ({
      element: reference.element,
      relId: `rId${next++}`,
      relType: reference.relType,
      entryPath: remap.get(reference.entryPath) as string,
    }));
  });

  const workbookRefs = workbook.preservedReferences.map((reference): PreservedWorkbookReferencePlan => ({
    relType: reference.relType,
    entryPath: remap.get(reference.entryPath) as string,
    pivotCacheId: reference.pivotCacheId,
  }));

  return {perSheet, workbook: workbookRefs, parts: [...emitted.values()]};
}

// Per-kind counters for numbering preserved parts, each seeded past the generated parts of its kind.
interface PreservedNumbering {
  drawing: number;
  vml: number;
  media: number;
}

// The path a preserved part is emitted at. A kind the writer generates of its own — a drawing, a VML,
// a media image — is re-numbered past the generated parts of that kind (see {@link planPreservedParts})
// so it never clobbers one. Every other kind (a pivot table, a pivot/slicer cache, a slicer, a chart)
// the writer never generates, so it keeps its original path — leaving the package's standard part
// names intact and letting overlapping closures agree on a single path for a shared part.
function preservedPartPath(originalPath: string, numbering: PreservedNumbering): string {
  const ext = extensionOf(originalPath);
  if (ext.toLowerCase() === 'vml') return `xl/drawings/vmlDrawing${++numbering.vml}.vml`;
  if (originalPath.startsWith('xl/media/')) return `xl/media/image${++numbering.media}.${ext}`;
  if (originalPath.startsWith('xl/drawings/') && ext.toLowerCase() === 'xml') {
    return `xl/drawings/drawing${++numbering.drawing}.xml`;
  }
  return originalPath;
}

// The extension of a part path (`xl/media/image1.jpeg` → `jpeg`), or '' when it carries none.
function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf('.');
  const slash = partPath.lastIndexOf('/');
  return dot > slash ? partPath.slice(dot + 1) : '';
}

// The relationships part path for `dir/name.ext` → `dir/_rels/name.ext.rels`.
function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const base = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// A relationship target expressed relative to the part that carries it: the `..` hops out of the
// referencing part's directory up to the common ancestor, then down to the target. Both paths are
// package-absolute (`xl/drawings/preservedP1.vml` → `xl/media/preservedP2.jpeg` → `../media/preservedP2.jpeg`).
function relativePartPath(fromPath: string, toPath: string): string {
  const fromDir = fromPath.split('/').slice(0, -1);
  const toSegments = toPath.split('/');
  let common = 0;
  while (common < fromDir.length && common < toSegments.length - 1 && fromDir[common] === toSegments[common]) {
    common++;
  }
  const up = fromDir.length - common;
  return [...Array<string>(up).fill('..'), ...toSegments.slice(common)].join('/');
}

function preservedRelsXml(rels: readonly {id: string; type: string; target: string}[]): string {
  const body = rels.map(rel => relationship(rel.id, rel.type, escapeAttr(rel.target))).join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${body}</Relationships>`;
}

// A `.rels` part for a generated part chain (pivot table → cache definition → cache records). Targets
// are writer-controlled package paths, so no attribute escaping is needed.
function relsPartXml(rels: readonly {id: string; type: string; target: string}[]): string {
  const body = rels.map(rel => relationship(rel.id, rel.type, rel.target)).join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${body}</Relationships>`;
}

function contentTypesXml(
  sheetCount: number,
  tables: readonly PlannedTable[],
  commentNumbers: readonly number[],
  drawingNumbers: readonly number[],
  printerSettingsNumbers: readonly number[],
  mediaExtensions: readonly string[],
  hasSharedStrings: boolean,
  preservedParts: readonly PreservedPartPlan[],
  pivots: readonly PivotPlan[]
): string {
  // A preserved part with its own XML content type (a drawing) needs an <Override>; a binary one (a
  // VML, an image) is declared by a <Default> for its extension, deduped against the defaults already
  // emitted (rels, xml, vml, bin, the media kinds) and against each other.
  const declaredExtensions = new Set<string>(['rels', 'xml']);
  if (commentNumbers.length > 0) declaredExtensions.add('vml');
  if (printerSettingsNumbers.length > 0) declaredExtensions.add('bin');
  for (const ext of mediaExtensions) declaredExtensions.add(ext.toLowerCase());
  const preservedOverrides: string[] = [];
  const preservedDefaults: string[] = [];
  for (const part of preservedParts) {
    const ext = part.path.slice(part.path.lastIndexOf('.') + 1);
    if (ext.toLowerCase() === 'xml') {
      preservedOverrides.push(override(`/${part.path}`, part.contentType));
    } else if (!declaredExtensions.has(ext.toLowerCase())) {
      declaredExtensions.add(ext.toLowerCase());
      preservedDefaults.push(`<Default Extension="${ext}" ContentType="${part.contentType}"/>`);
    }
  }

  const overrides = [
    override('/xl/workbook.xml', CT.workbook),
    ...range(sheetCount).map(i => override(`/xl/worksheets/sheet${i + 1}.xml`, CT.worksheet)),
    ...tables.map(({number}) => override(`/xl/tables/table${number}.xml`, CT.table)),
    ...drawingNumbers.map(number => override(`/xl/drawings/drawing${number}.xml`, CT.drawing)),
    ...commentNumbers.map(number => override(`/xl/comments${number}.xml`, CT.comments)),
    ...pivots.map(({number}) => override(`/xl/pivotTables/pivotTable${number}.xml`, CT.pivotTable)),
    ...pivots.map(({number}) =>
      override(`/xl/pivotCache/pivotCacheDefinition${number}.xml`, CT.pivotCacheDefinition)
    ),
    ...pivots.map(({number}) =>
      override(`/xl/pivotCache/pivotCacheRecords${number}.xml`, CT.pivotCacheRecords)
    ),
    override('/xl/theme/theme1.xml', CT.theme),
    override('/xl/styles.xml', CT.styles),
    ...(hasSharedStrings ? [override('/xl/sharedStrings.xml', CT.sharedStrings)] : []),
    override('/docProps/core.xml', CT.core),
    override('/docProps/app.xml', CT.app),
    ...preservedOverrides,
  ].join('');
  // The VML drawings, printer-settings blobs, and each media kind are declared by extension-level
  // defaults rather than a per-part override — the raw bytes carry no XML content type of their own.
  const vmlDefault =
    commentNumbers.length > 0 ? `<Default Extension="vml" ContentType="${CT.vml}"/>` : '';
  const binDefault =
    printerSettingsNumbers.length > 0 ? `<Default Extension="bin" ContentType="${CT.printerSettings}"/>` : '';
  const imageDefaults = mediaExtensions
    .map(ext => `<Default Extension="${ext}" ContentType="${imageContentType(ext)}"/>`)
    .join('');
  return (
    XML_DECLARATION +
    `<Types xmlns="${NS.contentTypes}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    vmlDefault +
    binDefault +
    imageDefaults +
    preservedDefaults.join('') +
    overrides +
    '</Types>'
  );
}

function override(partName: string, contentType: string): string {
  return `<Override PartName="${partName}" ContentType="${contentType}"/>`;
}

function rootRelsXml(): string {
  const rels = [
    relationship('rId1', REL.officeDocument, 'xl/workbook.xml'),
    relationship('rId2', REL.coreProps, 'docProps/core.xml'),
    relationship('rId3', REL.extProps, 'docProps/app.xml'),
  ].join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

function workbookXml(
  workbook: Workbook,
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[]
): string {
  const sheets = workbook.worksheets;
  const entries = sheets
    .map((sheet, i) => {
      const state = sheet.state === 'visible' ? '' : ` state="${sheet.state}"`;
      return `<sheet name="${escapeAttr(sheet.name)}" sheetId="${sheet.id}"${state} r:id="rId${i + 1}"/>`;
    })
    .join('');
  return (
    XML_DECLARATION +
    `<workbook xmlns="${NS.main}" xmlns:r="${NS.docRels}">` +
    `<sheets>${entries}</sheets>` +
    definedNamesXml(workbook) +
    calcPrXml(workbook) +
    pivotCachesXml(preservedRels, pivots) +
    workbookExtLstXml(preservedRels) +
    '</workbook>'
  );
}

// The workbook-body `<x14:slicerCaches>` extension that registers each preserved slicer cache, wired
// to the relationship reaching its cache part. Slicer caches (unlike pivot caches, which register in
// `<pivotCaches>`) live only in this extension block, so re-emitting it is what lets Excel rediscover
// the slicers. `<extLst>` is the final child of CT_Workbook. '' when no slicer cache was preserved.
function workbookExtLstXml(preservedRels: readonly PreservedWorkbookRel[]): string {
  const caches = preservedRels.filter(ref => ref.relType.endsWith('/slicerCache'));
  if (caches.length === 0) return '';
  const entries = caches.map(ref => `<x14:slicerCache r:id="${ref.relId}"/>`).join('');
  return (
    `<extLst><ext uri="${SLICER_CACHES_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:slicerCaches>${entries}</x14:slicerCaches></ext></extLst>`
  );
}

// The `<pivotCaches>` element registers each pivot cache under the `cacheId` a pivot table resolves
// its cache through, wired to the relationship that reaches the cache definition. It follows
// `<calcPr>` in CT_Workbook order and carries both preserved caches (passed through from a read file)
// and caches the writer generated for modeled pivot tables. A slicer cache (no `cacheId`) is
// registered in a workbook extension block, not here, so it is skipped.
function pivotCachesXml(
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[]
): string {
  const preserved = preservedRels
    .filter(ref => ref.pivotCacheId !== undefined)
    .map(ref => `<pivotCache cacheId="${escapeAttr(ref.pivotCacheId as string)}" r:id="${ref.relId}"/>`);
  const generated = pivots.map(
    pivot => `<pivotCache cacheId="${escapeAttr(pivot.cacheId)}" r:id="${pivot.workbookRelId}"/>`
  );
  const entries = [...preserved, ...generated];
  if (entries.length === 0) return '';
  return `<pivotCaches>${entries.join('')}</pivotCaches>`;
}

// `<calcPr>` follows `<definedNames>` in CT_Workbook order and carries the calculation settings.
// Today the model exposes a single one: `fullCalcOnLoad`, which tells the consumer to recalculate
// every formula on open instead of trusting the cached results. Emitted only when set, so an
// unmarked workbook keeps the element (and its `calcId`) out of the file entirely.
function calcPrXml(workbook: Workbook): string {
  return workbook.fullCalcOnLoad ? '<calcPr calcId="171027" fullCalcOnLoad="1"/>' : '';
}

// The `<definedNames>` block follows `<sheets>` in the schema. A sheet-scoped name carries a
// `localSheetId` — the 0-based position of its sheet among the `<sheet>` entries, NOT the sheet's
// own id — so the index is resolved against the worksheet order here. The refersTo formula is the
// element's text content, run through the same `_xlfn.` function mangling the writer applies to a
// cell formula so a name defined as a modern function (a LAMBDA, an XLOOKUP-based name) is stored
// under the prefix Excel requires; a plain reference has no function call and passes through
// untouched. Only names that are actually set emit anything.
function definedNamesXml(workbook: Workbook): string {
  const sheets = workbook.worksheets;
  const userEntries = workbook.definedNames.map(name => {
    const scopeAttr =
      name.scope === undefined
        ? ''
        : ` localSheetId="${sheets.findIndex(sheet => sheet.name === name.scope)}"`;
    const commentAttr = name.comment === undefined ? '' : ` comment="${escapeAttr(name.comment)}"`;
    const hiddenAttr = name.hidden ? ' hidden="1"' : '';
    return (
      `<definedName name="${escapeAttr(name.name)}"${scopeAttr}${commentAttr}${hiddenAttr}>` +
      `${escapeText(mangleFormula(name.refersTo))}</definedName>`
    );
  });
  // Every sheet-level autofilter contributes the hidden, sheet-scoped `_FilterDatabase` built-in that
  // Excel derives from its range. The reader drops these on load and rebuilds them from the sheet's
  // `<autoFilter>`, so `Worksheet.autoFilter` stays the single source of truth and a round-trip never
  // duplicates them.
  const filterEntries = sheets.flatMap((sheet, index) =>
    sheet.autoFilter === undefined
      ? []
      : [
          `<definedName name="_xlnm._FilterDatabase" localSheetId="${index}" hidden="1">` +
            `${escapeText(filterDatabaseRefersTo(sheet.name, sheet.autoFilter.ref))}</definedName>`,
        ]
  );

  const entries = [...userEntries, ...filterEntries];
  if (entries.length === 0) return '';
  return `<definedNames>${entries.join('')}</definedNames>`;
}

// Build the sheet-qualified, fully-absolute reference a `_FilterDatabase` name carries
// (`'Sheet 1'!$A$1:$C$10`) from a sheet name and its already-canonical `A1:C10` autofilter range.
function filterDatabaseRefersTo(sheetName: string, range: string): string {
  const absolute = range.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  return `${quoteSheetName(sheetName)}!${absolute}`;
}

// Quote a sheet name for use in a reference exactly when Excel would: a name that is not a plain
// identifier (or that looks like a cell address) is wrapped in single quotes with internal quotes
// doubled; a simple name is left bare so the output matches what Excel writes.
function quoteSheetName(name: string): string {
  const bare = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name);
  return bare ? name : `'${name.replace(/'/g, "''")}'`;
}

function workbookRelsXml(
  sheetCount: number,
  hasSharedStrings: boolean,
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[]
): string {
  const rels = [
    ...range(sheetCount).map(i =>
      relationship(`rId${i + 1}`, REL.worksheet, `worksheets/sheet${i + 1}.xml`)
    ),
    relationship(`rId${sheetCount + 1}`, REL.styles, 'styles.xml'),
    relationship(`rId${sheetCount + 2}`, REL.theme, 'theme/theme1.xml'),
    ...(hasSharedStrings
      ? [relationship(`rId${sheetCount + 3}`, REL.sharedStrings, 'sharedStrings.xml')]
      : []),
    // A preserved cache's target is package-absolute; express it relative to the workbook part.
    ...preservedRels.map(ref =>
      relationship(ref.relId, ref.relType, escapeAttr(relativePartPath('xl/workbook.xml', ref.entryPath)))
    ),
    // A generated pivot cache's workbook relationship reaches its cache definition part.
    ...pivots.map(pivot =>
      relationship(
        pivot.workbookRelId,
        REL.pivotCacheDefinition,
        `pivotCache/pivotCacheDefinition${pivot.number}.xml`
      )
    ),
  ].join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

// A preserved workbook reference with the relationship id assigned for emission (see the body and
// rels-part wiring in {@link buildPackageParts}).
type PreservedWorkbookRel = PreservedWorkbookReferencePlan & {readonly relId: string};

function relationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
}

function corePropsXml(properties: WorkbookProperties): string {
  const parts: string[] = [];
  if (properties.creator !== undefined) {
    parts.push(`<dc:creator>${escapeText(properties.creator)}</dc:creator>`);
  }
  if (properties.lastModifiedBy !== undefined) {
    parts.push(`<cp:lastModifiedBy>${escapeText(properties.lastModifiedBy)}</cp:lastModifiedBy>`);
  }
  if (properties.created) {
    parts.push(w3cdtf('created', properties.created));
  }
  if (properties.modified) {
    parts.push(w3cdtf('modified', properties.modified));
  }
  return (
    XML_DECLARATION +
    `<cp:coreProperties xmlns:cp="${NS.coreProps}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" ` +
    `xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">` +
    parts.join('') +
    '</cp:coreProperties>'
  );
}

function w3cdtf(element: string, date: Date): string {
  return `<dcterms:${element} xsi:type="dcterms:W3CDTF">${date.toISOString()}</dcterms:${element}>`;
}

function appPropsXml(): string {
  return (
    XML_DECLARATION +
    `<Properties xmlns="${NS.extProps}"><Application>ts-xlsx</Application></Properties>`
  );
}

function worksheetXml(
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
  sharedStrings: SharedStringTable | null
): string {
  // A merge overlapping a table is Excel-invalid geometry; reject it before serialising
  // rather than emit a package a consumer repairs on open.
  validateMerges(sheet);

  // A column's style facets are defaults its cells inherit unless they override them; the writer
  // composes each cell's full style up front (cell over row over column, per facet) so a cell that
  // overrides one facet still carries the column's others, rather than silently dropping them.
  const columnDefaults = new Map<number, ColumnProperties>();
  for (const {index, properties} of sheet.columns()) columnDefaults.set(index, properties);

  // A cell filled from a shared formula is written as a master (seeding the group) or a clone
  // (referencing it by shared index); resolve every such role before the row loop so each cell knows
  // how to serialise its `<f>`. This also validates the master/clone geometry, throwing if a clone
  // precedes its master or its master carries no formula.
  const sharedRoles = planSharedFormulas(sheet);

  const rowXml: string[] = [];
  let top = Infinity;
  let left = Infinity;
  let bottom = -Infinity;
  let right = -Infinity;

  for (const {number, cells, properties} of sheet.rows()) {
    // A cell earns a <c> element if it holds a value OR carries its own style: a formatted-but-empty
    // cell (a fill/border on a null value) is a real cell to Excel, and dropping it would lose the
    // formatting. A cell with neither is inherited from its row/column and needs no element of its own.
    const rendered = cells.filter(cell => cell.value !== null || hasOwnStyle(cell));
    const attrs = rowAttrs(properties, styles);
    // A row with neither data nor its own formatting has nothing to serialise.
    if (rendered.length === 0 && attrs === '') continue;
    const rowFill = properties?.fill;
    const cellsXml = rendered
      .map(cell => {
        // Cell overrides win over row/column defaults; a cell with any facet gets its own,
        // fully-composed style entry so no default facet is lost to the override. Precedence is
        // cell over row over column per facet — the row contributes only a fill today.
        const colDef = columnDefaults.get(cell.col);
        const style = styles.styleId({
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
        return cellXml(cell, style, sharedRoles.get(cell.address), sharedStrings);
      })
      .join('');
    rowXml.push(`<row r="${number}"${attrs}>${cellsXml}</row>`);
    for (const cell of rendered) {
      if (number < top) top = number;
      if (number > bottom) bottom = number;
      if (cell.col < left) left = cell.col;
      if (cell.col > right) right = cell.col;
    }
  }

  // Dimension is the used *cell* range; rows/columns carrying only formatting do not
  // extend it, matching how Excel records <dimension>.
  const dimensionRef =
    bottom === -Infinity ? 'A1' : `${encodeAddress(left, top)}:${encodeAddress(right, bottom)}`;
  const sheetData = rowXml.length === 0 ? '<sheetData/>' : `<sheetData>${rowXml.join('')}</sheetData>`;

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
    pageMarginsXml(sheet.pageMargins) +
    pageSetupXml(sheet.pageSetup, printerSettingsRelId) +
    headerFooterXml(sheet.headerFooter) +
    // CT_Worksheet order: <rowBreaks> follows <headerFooter> and precedes the drawing block.
    rowBreaksXml(sheet.rowBreaks) +
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

// Assemble the worksheet's single `<extLst>` from every x14 extension the sheet carries, or '' when it
// carries none. Each producer returns a bare `<ext>` so they compose without nesting an `<extLst>`.
function worksheetExtLstXml(sheet: Worksheet, slicerRelIds: readonly string[]): string {
  const exts = [
    conditionalFormattingsExtXml(sheet.conditionalFormattings),
    dataValidationsExtXml(sheet.dataValidations),
    slicerListExtXml(slicerRelIds),
  ].filter(ext => ext !== '');
  return exts.length === 0 ? '' : `<extLst>${exts.join('')}</extLst>`;
}

// The x14 namespace and the well-known extension URIs Excel keys the slicer wiring off. The `<ext>`
// blocks are opaque to any consumer that does not understand them, so a producer must reproduce these
// exact GUIDs for Excel to rediscover the slicer widget and its caches.
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const SLICER_LIST_EXT_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}';
const SLICER_CACHES_EXT_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}';

// The worksheet-body `<x14:slicerList>` extension that reconnects a sheet to its preserved slicer
// parts. Each `<x14:slicer>` names the sheet-local relationship id its slicer rel was re-emitted under,
// so the wiring stays consistent even though the id is reassigned on write. '' when the sheet has none.
function slicerListExtXml(slicerRelIds: readonly string[]): string {
  if (slicerRelIds.length === 0) return '';
  const slicers = slicerRelIds.map(relId => `<x14:slicer r:id="${relId}"/>`).join('');
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
    if (left === undefined || right === undefined || top === undefined || bottom === undefined) continue;
    for (const table of sheet.tables) {
      const region = table.region;
      const overlaps =
        left <= region.right && right >= region.left && top <= region.bottom && bottom >= region.top;
      if (overlaps) {
        throw new Error(
          `merged range ${merge} overlaps table "${table.name}" (${table.ref}) — Excel forbids a merge inside a table`
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
  const activePane = xSplit > 0 && ySplit > 0 ? 'bottomRight' : xSplit > 0 ? 'topRight' : 'bottomLeft';
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
  if (outline.summaryBelow !== undefined) attrs.push(`summaryBelow="${outline.summaryBelow ? 1 : 0}"`);
  if (outline.summaryRight !== undefined) attrs.push(`summaryRight="${outline.summaryRight ? 1 : 0}"`);
  return attrs.length === 0 ? '' : `<outlinePr ${attrs.join(' ')}/>`;
}

function mergeCellsXml(merges: readonly string[]): string {
  if (merges.length === 0) return '';
  const cells = merges.map(range => `<mergeCell ref="${escapeAttr(decodeRange(range).dimensions)}"/>`).join('');
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
    const filters = criteria.values.map(value => `<filter val="${escapeAttr(value)}"/>`).join('');
    return `<filters${blankAttr}>${filters}</filters>`;
  }
  const andAttr = criteria.and ? ' and="1"' : '';
  const predicates = criteria.predicates
    .map(p => `<customFilter operator="${p.operator}" val="${escapeAttr(p.val)}"/>`)
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

function worksheetRelsXml(
  tables: readonly PlannedTable[],
  drawing: DrawingPlan | null,
  comments: CommentPlan | null,
  printerSettings: PrinterSettingsPlan | null,
  background: BackgroundPlan | null,
  hyperlinks: readonly PlannedHyperlink[],
  preservedReferences: readonly PreservedReferencePlan[],
  pivots: readonly PivotPlan[]
): string {
  const rels = [
    ...tables.map(({relId, number}) => relationship(relId, REL.table, `../tables/table${number}.xml`)),
    // A pivot table hosted on this sheet is reached by a relationship of type pivotTable; Excel
    // discovers the pivot from the rels part, so the sheet body itself carries no reference to it.
    ...pivots.map(pivot =>
      relationship(pivot.sheetRelId, REL.pivotTable, `../pivotTables/pivotTable${pivot.number}.xml`)
    ),
    ...(drawing === null
      ? []
      : [relationship(drawing.relId, REL.drawing, `../drawings/drawing${drawing.number}.xml`)]),
    ...(comments === null
      ? []
      : [
          relationship(comments.vmlRelId, REL.vmlDrawing, `../drawings/vmlDrawing${comments.number}.vml`),
          relationship(comments.commentsRelId, REL.comments, `../comments${comments.number}.xml`),
        ]),
    ...(printerSettings === null
      ? []
      : [
          relationship(
            printerSettings.relId,
            REL.printerSettings,
            `../printerSettings/printerSettings${printerSettings.number}.bin`
          ),
        ]),
    ...(background === null
      ? []
      : [
          relationship(
            background.relId,
            REL.image,
            `../media/image${background.mediaNumber}.${background.extension}`
          ),
        ]),
    // A preserved reference targets its entry part's new (package-absolute) path; a worksheet always
    // lives under `xl/worksheets/`, so the target is that path made relative to that directory.
    ...preservedReferences.map(reference =>
      relationship(
        reference.relId,
        reference.relType,
        escapeAttr(relativePartPath('xl/worksheets/sheet1.xml', reference.entryPath))
      )
    ),
    // An external hyperlink's target is a URL outside the package, so its relationship carries
    // TargetMode="External" and the plain `relationship()` helper (a package-internal target) will
    // not do. Internal links have no relId and contribute nothing here.
    ...hyperlinks
      .filter((link) => link.relId !== undefined && link.target !== undefined)
      .map(
        (link) =>
          `<Relationship Id="${link.relId}" Type="${REL.hyperlink}" ` +
          `Target="${escapeAttr(link.target as string)}" TargetMode="External"/>`
      ),
  ].join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

function tableXml(table: Table, id: number): string {
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
  if (style.showFirstColumn !== undefined) attrs += ` showFirstColumn="${style.showFirstColumn ? '1' : '0'}"`;
  if (style.showLastColumn !== undefined) attrs += ` showLastColumn="${style.showLastColumn ? '1' : '0'}"`;
  if (style.showRowStripes !== undefined) attrs += ` showRowStripes="${style.showRowStripes ? '1' : '0'}"`;
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
  const body = children.map(({tag, key}) => `<${tag}>${escapeText(hf[key] as string)}</${tag}>`).join('');
  return `<headerFooter${attrs}>${body}</headerFooter>`;
}

// Excel's "Normal" margins, in inches — the defaults Excel writes for an untouched sheet.
const DEFAULT_MARGINS = {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3} as const;
const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

// OOXML's <pageMargins> is all-or-nothing: setting any one margin requires all six, or Excel
// repairs the file. So the element is emitted only when the caller set at least one, and the
// untouched sides fall back to the Normal-preset defaults.
function pageMarginsXml(margins: PageMargins): string {
  if (MARGIN_SIDES.every(side => margins[side] === undefined)) return '';
  const attrs = MARGIN_SIDES.map(
    side => `${side}="${numberText(margins[side] ?? DEFAULT_MARGINS[side])}"`
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
  const cols = runs.map(run => `<col min="${run.min}" max="${run.max}"${run.body}/>`).join('');
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

// Manual horizontal page breaks (`<rowBreaks>`): one `<brk>` per row the layout splits before. Excel
// records both the running total (`count`) and the manual subset (`manualBreakCount`); every break the
// model carries is a manual, author-set one, so the two counts coincide. `max` bounds the break across
// columns (Excel writes the last column index); a break without one is emitted bare.
function rowBreaksXml(breaks: readonly PageBreak[]): string {
  if (breaks.length === 0) return '';
  const brks = breaks
    .map(brk => {
      const maxAttr = brk.max !== undefined ? ` max="${brk.max}"` : '';
      return `<brk id="${brk.id}"${maxAttr} man="1"/>`;
    })
    .join('');
  return `<rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">${brks}</rowBreaks>`;
}

function rowAttrs(properties: RowProperties | undefined, styles: StyleRegistry): string {
  if (properties === undefined) return '';
  let attrs = '';
  if (properties.height !== undefined) attrs += ` ht="${numberText(properties.height)}" customHeight="1"`;
  if (properties.hidden) attrs += ' hidden="1"';
  if (properties.outlineLevel !== undefined && properties.outlineLevel > 0) {
    attrs += ` outlineLevel="${properties.outlineLevel}"`;
  }
  if (properties.collapsed) attrs += ' collapsed="1"';
  // A row-level fill is a default format for the row's cells; customFormat="1" is what makes
  // Excel honour the row's `s`, and a cell without its own `s` then inherits it.
  const style = styles.styleId({fill: properties.fill});
  if (style !== 0) attrs += ` s="${style}" customFormat="1"`;
  return attrs;
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
  sharedStrings: SharedStringTable | null
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
    const label = typeof value.text === 'string' ? textElement(value.text) : richTextRunsXml(value.text.richText);
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

function formulaCellXml(ref: string, s: string, formula: string, result: FormulaResult | undefined): string {
  return formulaBodyXml(ref, s, `<f>${escapeText(mangleFormula(formula))}</f>`, result);
}

// Wrap a prepared `<f>` element (a plain formula, or a shared master/slave `<f>`) with the cell
// element and its cached result, typing the cell by the result's kind exactly as a bare value of that
// kind would be.
function formulaBodyXml(ref: string, s: string, f: string, result: FormulaResult | undefined): string {
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
interface SharedFormulaRole {
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
        `shared-formula clone ${offender.address} names master ${masterAddress}, which holds no formula`
      );
    }
    let maxCol = master.col;
    let maxRow = master.row;
    for (const clone of clones) {
      if (clone.col < master.col || clone.row < master.row) {
        throw new Error(
          `shared-formula master ${masterAddress} must sit above and/or left of clone ${clone.address}`
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

// A finite number serialises as its shortest round-trippable decimal; a non-finite one
// has no OOXML numeric representation, so the writer refuses it rather than emit `NaN`.
function numberText(value: number): string {
  if (!Number.isFinite(value)) {
    throw new Error(`cannot write a non-finite number (${value}) — it has no OOXML representation`);
  }
  return String(value);
}

function range(n: number): number[] {
  return Array.from({length: n}, (_, i) => i);
}
