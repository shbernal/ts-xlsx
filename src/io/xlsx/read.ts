// The buffered `.xlsx` reader: an OPC zip package in, a Workbook model out.
//
// It reconstructs the part of the model the writer emits today — sheet names and order,
// cells holding a number, string, boolean, or formula, per-column width/visibility,
// per-row height/visibility, merged ranges, page margins, and cell styles (pattern fills,
// number formats, fonts, borders, alignment, and protection — per cell, or inherited from a
// formatted row/column). Shared-formula slaves and the richer value kinds land as the model
// grows; an unrecognised construct is skipped rather than guessed, so a foreign file reads
// without crashing even where a facet is not yet materialised.
//
// This module is the orchestrator: it wires the parsed package parts together — the OPC/rel
// resolution (`./read-opc.ts`), the style table (`./read-styles.ts`), and each worksheet body
// (`./read-worksheet.ts`) — and owns the sheet-part discovery (notes, images, tables, pivots) and
// preserved-reference capture that a faithful round-trip depends on.
//
// Untrusted input: inflate is bounded by a running byte counter (`./inflate.ts`) that caps
// actual decompressed output rather than trusting the archive's forgeable size headers, and
// the parser (ADR 0004) never expands entities.

import {strFromU8} from 'fflate';

import {decodeRange} from '../../core/address.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import type {PreservedWorksheetReference} from '../../core/preserved.ts';
import {type DefinedName, Workbook} from '../../core/workbook.ts';
import {
  WORKBOOK_PROTECTION_CREDENTIAL_ATTRS,
  type WorkbookProtection,
} from '../../core/workbook-protection.ts';
import type {Worksheet, WorksheetState} from '../../core/worksheet.ts';
import type {SharedString} from './cell-value.ts';
import {applyNotes, parseComments} from './comments.ts';
import {parseConditionalFormattings, parseDxfs} from './conditional-formatting.ts';
import {
  applyDataValidations,
  parseDataValidations,
  parseExtendedDataValidations,
} from './data-validation.ts';
import {applyHyperlinks, parseSheetHyperlinks} from './hyperlinks.ts';
import {drawingHasUnmodeledContent, parseDrawing} from './images.ts';
import {inflatePackage} from './inflate.ts';
import {parsePivotTable} from './pivot-read.ts';
import {
  capturePartClosure,
  contentTypeResolver,
  extensionOf,
  parseRelationshipRecords,
  parseRelationships,
  relationshipTargetByType,
  relationshipTargetsByType,
  relsPathFor,
  resolveRelativePart,
  resolveWorkbookPart,
} from './read-opc.ts';
import {parseStyleTable} from './read-styles.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {RunAccumulator} from './rich-runs.ts';
import {parseIndexedColors} from './styles.ts';
import {parseTable} from './tables.ts';
import {boolStrict, localName, openElements, parseXml} from './xml-read.ts';

// Re-exported for the streaming reader (`./read-rows.ts`) and the public barrel, which import these
// from here: the split into per-part parsers is internal, so the reader's import surface is stable.
export {parseRelationships, resolveWorkbookPart} from './read-opc.ts';
export {parseStyleTable, type StyleTable, type XfStyle} from './read-styles.ts';

export interface ReadXlsxOptions {
  /**
   * Maximum total uncompressed output, in bytes, produced while inflating the package.
   * The bound is enforced by a running counter as bytes are decompressed — never read from
   * the archive's (untrusted, forgeable) size headers — so a zip bomb that lies about its
   * uncompressed size is rejected all the same. Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}

export const DEFAULT_MAX_UNCOMPRESSED = 512 * 1024 * 1024;

/**
 * Read an `.xlsx` package into a {@link Workbook}.
 *
 * @throws {Error} if the archive is malformed, exceeds the inflate bound, or names no
 *   worksheet parts (a workbook with no sheets is not a valid package).
 */
export function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook {
  const cap = options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  const files = inflatePackage(data, cap);

  const partText = (path: string): string | undefined => {
    const bytes = files[path];
    return bytes === undefined ? undefined : strFromU8(bytes);
  };
  const partBytes = (path: string): Uint8Array | undefined => files[path];

  const workbookXml = partText('xl/workbook.xml');
  if (workbookXml === undefined) throw new Error('not an xlsx package: xl/workbook.xml is missing');

  // A part's content type is needed to faithfully re-declare any part preserved verbatim for
  // round-tripping (a vector-shape drawing, a header/footer image and its VML). Resolve it the way
  // OPC does: an explicit `<Override>` for the exact part, else the `<Default>` for its extension.
  const contentTypeOf = contentTypeResolver(partText('[Content_Types].xml') ?? '');

  const rels = parseRelationships(partText('xl/_rels/workbook.xml.rels') ?? '');
  const sharedStrings = parseSharedStrings(partText('xl/sharedStrings.xml') ?? '');
  // The style table resolves a cell/row/column style index to its facets (fill, number
  // format); a package without one (a hand-rolled foreign file) yields an empty table and
  // every index reads as unstyled.
  const stylesXml = partText('xl/styles.xml') ?? '';
  const {cellXfs: xfStyles, namedStyles} = parseStyleTable(stylesXml);

  const workbook = new Workbook();
  // Preserve the differential-style table verbatim so conditional formatting's dxfId references stay
  // valid — and a foreign dxf's number format stays a real format code — across a re-write.
  workbook.restoreDifferentialStyles(parseDxfs(stylesXml));
  // Preserve a custom indexed-color palette verbatim so an `indexed="…"` colour reference keeps its
  // intended RGB across a re-write instead of resolving to a different default-palette entry.
  workbook.restoreIndexedColors(parseIndexedColors(stylesXml));
  // Preserve the named cell-style layer only when a file declares one beyond the Normal default, so an
  // ordinary workbook keeps an empty named-style table and emits just the default on write.
  if (namedStyles.length > 1) workbook.restoreNamedStyles(namedStyles);
  const core = partText('docProps/core.xml');
  if (core !== undefined) applyCoreProperties(workbook, core);
  workbook.protection = parseWorkbookProtection(workbookXml);

  // A picture used on more than one sheet is one media part; caching by media path keeps it a single
  // workbook image so a re-write does not duplicate the bytes.
  const imageIdByMediaPath = new Map<string, number>();
  const sheetOrder: string[] = [];
  for (const {name, relId, state} of parseWorkbookSheets(workbookXml)) {
    const target = rels.get(relId);
    const sheet = workbook.addWorksheet(name, state === undefined ? undefined : {state});
    sheetOrder.push(name);
    const path = target === undefined ? undefined : resolveWorkbookPart(target);
    const sheetXml = path === undefined ? undefined : partText(path);
    if (sheetXml !== undefined) parseWorksheet(sheetXml, sheet, sharedStrings, xfStyles);
    if (path !== undefined) {
      if (sheetXml !== undefined) {
        const sheetRels = parseRelationships(partText(relsPathFor(path)) ?? '');
        applyHyperlinks(sheet, parseSheetHyperlinks(sheetXml), sheetRels);
        applyDataValidations(sheet, [
          ...parseDataValidations(sheetXml),
          ...parseExtendedDataValidations(sheetXml),
        ]);
        for (const cf of parseConditionalFormattings(sheetXml)) sheet.addConditionalFormatting(cf);
      }
      const notes = readSheetNotes(path, partText);
      if (notes !== undefined) applyNotes(sheet, notes);
      readSheetImages(path, partText, partBytes, workbook, sheet, imageIdByMediaPath);
      readSheetBackground(path, partText, partBytes, workbook, sheet, imageIdByMediaPath);
      if (sheetXml !== undefined) {
        readSheetPreservedReferences(path, sheetXml, partText, partBytes, contentTypeOf, sheet);
      }
      readSheetTables(path, partText, sheet);
      readSheetPivotTables(path, partText, sheet);
      const printerSettings = readSheetPrinterSettings(path, partText, partBytes);
      if (printerSettings !== undefined) sheet.pageSetup.printerSettings = printerSettings;
    }
  }

  readWorkbookPreservedReferences(workbookXml, partText, partBytes, contentTypeOf, workbook);

  // Defined names follow the sheets: a scoped name's `localSheetId` indexes the sheet order, which
  // is why the names are read only once every sheet is registered.
  for (const name of parseWorkbookDefinedNames(workbookXml, sheetOrder)) {
    workbook.defineName(name);
  }
  return workbook;
}

// A sheet's notes live in a comments part reached through the sheet's own relationships: the sheet
// declares a relationship of type `.../comments` whose target resolves (relative to the sheet's
// directory) to the comments part. A sheet with no rels part or no such relationship simply has none.
function readSheetNotes(
  sheetPath: string,
  partText: (path: string) => string | undefined,
): Map<string, string> | undefined {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return undefined;
  const target = relationshipTargetByType(relsXml, 'comments');
  if (target === undefined) return undefined;
  const commentsXml = partText(resolveRelativePart(sheetPath, target));
  if (commentsXml === undefined) return undefined;
  return parseComments(commentsXml);
}

// A sheet's printer-settings blob is an opaque binary part linked from `<pageSetup r:id>`: the sheet
// declares a relationship of type `.../printerSettings` whose target resolves to a `.bin` part. We
// keep the raw bytes verbatim — the DEVMODE inside is platform-specific and the model never
// interprets it, only round-trips it so re-writing the file preserves the user's print configuration.
// A sheet with no rels part or no such relationship simply has none.
function readSheetPrinterSettings(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
): Uint8Array | undefined {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return undefined;
  const target = relationshipTargetByType(relsXml, 'printerSettings');
  if (target === undefined) return undefined;
  return partBytes(resolveRelativePart(sheetPath, target));
}

// A sheet's anchored images live in a drawing part reached through the sheet's own relationships: a
// relationship of type `.../drawing` names the drawing part, whose own relationships map each
// picture's embed id to a media part under `xl/media/`. Each anchor becomes a workbook image (deduped
// by media path) placed back on the sheet at its two-cell anchor.
function readSheetImages(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  const drawingTarget = relationshipTargetByType(relsXml, 'drawing');
  if (drawingTarget === undefined) return;
  const drawingPath = resolveRelativePart(sheetPath, drawingTarget);
  const drawingXml = partText(drawingPath);
  if (drawingXml === undefined) return;
  // A drawing that also holds a chart or shape is preserved whole (see readSheetPreservedReferences),
  // so its pictures must not be modeled here: modeling them would leave the sheet with images, which
  // suppresses that preservation and drops the chart. Leaving `sheet.images` empty routes the entire
  // drawing — pictures included — through byte-preservation, keeping every anchor faithful.
  if (drawingHasUnmodeledContent(drawingXml)) return;
  const drawingRels = parseRelationships(partText(relsPathFor(drawingPath)) ?? '');

  for (const anchor of parseDrawing(drawingXml)) {
    const target = drawingRels.get(anchor.embed);
    if (target === undefined) continue;
    const mediaPath = resolveRelativePart(drawingPath, target);
    let id = imageIdByMediaPath.get(mediaPath);
    if (id === undefined) {
      const bytes = partBytes(mediaPath);
      if (bytes === undefined) continue;
      id = workbook.addImage({buffer: bytes, extension: extensionOf(mediaPath)});
      imageIdByMediaPath.set(mediaPath, id);
    }
    const rot = anchor.rotation !== undefined ? {rotation: anchor.rotation} : {};
    if (anchor.to !== undefined) {
      const mode = anchor.editAs !== undefined ? {editAs: anchor.editAs} : {};
      sheet.addImageAnchor(id, {from: anchor.from, to: anchor.to, ...mode, ...rot});
    } else if (anchor.ext !== undefined) {
      sheet.addImageAnchor(id, {from: anchor.from, ext: anchor.ext, ...rot});
    }
  }
}

// A sheet background is a workbook image referenced by the worksheet's `<picture>` element through a
// sheet-local relationship of type `.../image`. Unlike an anchored image (whose image relationships
// live in the drawing part's own rels), the background's relationship sits directly on the sheet, so
// it is the sheet rels' sole image relationship. The bytes are deduped against images shared with a
// drawing, keeping one media part per picture across a re-write.
function readSheetBackground(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  const target = relationshipTargetByType(relsXml, 'image');
  if (target === undefined) return;
  const mediaPath = resolveRelativePart(sheetPath, target);
  let id = imageIdByMediaPath.get(mediaPath);
  if (id === undefined) {
    const bytes = partBytes(mediaPath);
    if (bytes === undefined) return;
    id = workbook.addImage({buffer: bytes, extension: extensionOf(mediaPath)});
    imageIdByMediaPath.set(mediaPath, id);
  }
  sheet.addBackgroundImage(id);
}

// Capture the worksheet-level references to package content the model does not interpret, so a
// round-trip re-emits them verbatim instead of dropping them:
//   • `<drawing>` — but only when the reader modeled no anchored image from it: either a drawing that
//     holds no pictures at all (a chart or shape), or a mixed drawing whose pictures the reader
//     declined to model precisely so the whole part (chart included) rides here verbatim. A drawing
//     whose pictures were modeled is owned by the model and re-serialised from it; capturing it here
//     too would double-emit those pictures.
//   • `<legacyDrawingHF>` — a header/footer image's VML, which the model never interprets.
// Each reference's target part and the transitive closure of parts it reaches (a VML's image, a
// drawing's media) are captured with their bytes, content types, and relationships.
function readSheetPreservedReferences(
  sheetPath: string,
  sheetXml: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  contentTypeOf: (path: string) => string,
  sheet: Worksheet,
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  const records = parseRelationshipRecords(relsXml);
  const recordById = new Map(records.map((record) => [record.id, record]));

  const capture = (
    element: PreservedWorksheetReference['element'],
    relType: string,
    target: string,
  ): void => {
    const entryPath = resolveRelativePart(sheetPath, target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts !== undefined) sheet.addPreservedReference({element, relType, entryPath, parts});
  };

  // Element-wired references: a `<drawing>`/`<legacyDrawingHF>` names its part by an `r:id` in the
  // sheet body. A `<drawing>` is preserved only when the reader modeled no picture from it — a
  // chart/shape-only drawing, or a mixed one the reader left unmodeled — since one whose pictures are
  // modeled is re-serialised from the model.
  const referenceElements: Array<'drawing' | 'legacyDrawingHF'> =
    sheet.images.length === 0 ? ['drawing', 'legacyDrawingHF'] : ['legacyDrawingHF'];
  for (const element of referenceElements) {
    const relId = worksheetReferenceRelId(sheetXml, element);
    const record = relId === undefined ? undefined : recordById.get(relId);
    if (record !== undefined && !record.external) capture(element, record.type, record.target);
  }

  // Relationship-wired references: a pivot table or slicer is reached through a sheet relationship
  // with no worksheet child pointing at it — Excel discovers it by scanning the sheet's rels. Preserve
  // each so the pivots/slicers a fill-and-save workflow does not touch are not dropped.
  for (const record of records) {
    if (record.external) continue;
    if (isPreservedSheetRelType(record.type)) capture(null, record.type, record.target);
  }
}

// A sheet relationship the model does not consume but must round-trip: a pivot table or a slicer.
// The other sheet rel kinds (drawing, printerSettings, table, comments, hyperlinks, background image,
// the comment VML) are modeled and re-serialised from the model, so they are not preserved here.
function isPreservedSheetRelType(type: string): boolean {
  return type.endsWith('/pivotTable') || type.endsWith('/slicer');
}

// Capture the workbook-level references to package content the model does not interpret — pivot
// caches (`pivotCacheDefinition`) and slicer caches (`slicerCache`) — so a round-trip re-emits them
// instead of dropping the pivots and slicers they back. A pivot cache's `<pivotCaches>` registration
// (its `cacheId`) is captured alongside so the wiring a pivot table resolves its cache through
// survives too.
function readWorkbookPreservedReferences(
  workbookXml: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const relsXml = partText('xl/_rels/workbook.xml.rels');
  if (relsXml === undefined) return;
  const cacheIdByRelId = parsePivotCacheRegistrations(workbookXml);
  for (const record of parseRelationshipRecords(relsXml)) {
    if (record.external || !isPreservedWorkbookRelType(record.type)) continue;
    const entryPath = resolveWorkbookPart(record.target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts === undefined) continue;
    const cacheId = cacheIdByRelId.get(record.id);
    workbook.addPreservedReference({
      relType: record.type,
      entryPath,
      parts,
      ...(cacheId !== undefined ? {pivotCacheId: cacheId} : {}),
    });
  }
}

// A workbook relationship the model does not consume but must round-trip: a pivot cache or a slicer
// cache. Worksheets, styles, theme, and shared strings are modeled and re-serialised from the model.
function isPreservedWorkbookRelType(type: string): boolean {
  return type.endsWith('/pivotCacheDefinition') || type.endsWith('/slicerCache');
}

// Map each `<pivotCache>` registration in the workbook's `<pivotCaches>` to the relationship id that
// reaches its cache definition, so a preserved cache carries the `cacheId` a pivot table refers to.
function parsePivotCacheRegistrations(workbookXml: string): Map<string, string> {
  const byRelId = new Map<string, string>();
  for (const {attrs} of openElements(workbookXml, 'pivotCache')) {
    if (attrs['r:id'] !== undefined && attrs.cacheId !== undefined) {
      byRelId.set(attrs['r:id'], attrs.cacheId);
    }
  }
  return byRelId;
}

// The `r:id` of the first `<drawing>` / `<legacyDrawingHF>` element in a worksheet, or undefined when
// the sheet declares none. The reference lives in the worksheet XML (not distinguishable by
// relationship Type — a header/footer VML and a comment VML share the `vmlDrawing` type), so the
// specific relationship is found by reading the element's `r:id` here.
function worksheetReferenceRelId(sheetXml: string, element: string): string | undefined {
  for (const {attrs} of openElements(sheetXml, element)) {
    if (attrs['r:id'] !== undefined) return attrs['r:id'];
  }
  return undefined;
}

// A sheet's tables live in `xl/tables/table{n}.xml` parts, each reached through a relationship of
// type `.../table` on the sheet's own rels. The writer emits one relationship per table; each part
// is parsed back into the model and re-registered in definition order. A part that fails to parse
// (missing name/ref/columns — Excel corruption) is skipped rather than crashing the whole read.
function readSheetTables(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  sheet: Worksheet,
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  for (const target of relationshipTargetsByType(relsXml, 'table')) {
    const tableXml = partText(resolveRelativePart(sheetPath, target));
    if (tableXml === undefined) continue;
    const options = parseTable(tableXml);
    if (options !== undefined) sheet.addTable(options);
  }
  dropMergesInsideTables(sheet);
}

// Reconstruct an inspectable model of each pivot table hosted on a sheet. A pivot is reached by a
// sheet relationship of type `.../pivotTable`; the pivot-table part carries its own relationship of
// type `.../pivotCacheDefinition` to the cache holding the field catalogue and source range. Both
// parts are parsed and combined into a read-only view registered on the sheet — separate from the
// byte-preservation that actually round-trips the pivot, so this never changes what is re-emitted.
// The read is lenient: a pivot whose cache is missing still yields a (partial) model rather than
// throwing, matching Excel's tolerance for a damaged package on load.
function readSheetPivotTables(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  sheet: Worksheet,
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  for (const target of relationshipTargetsByType(relsXml, 'pivotTable')) {
    const tablePath = resolveRelativePart(sheetPath, target);
    const tableXml = partText(tablePath);
    if (tableXml === undefined) continue;
    const cacheTarget = relationshipTargetByType(
      partText(relsPathFor(tablePath)) ?? '',
      'pivotCacheDefinition',
    );
    const cacheXml =
      cacheTarget === undefined
        ? ''
        : (partText(resolveRelativePart(tablePath, cacheTarget)) ?? '');
    sheet.addLoadedPivotTable(parsePivotTable(tableXml, cacheXml));
  }
}

// Excel forbids a merged range inside a formatted table and repairs such a file on load by dropping
// the merge. A worksheet's merges are read before its tables, so a real file carrying that invalid
// geometry lands in the model intact; this applies the same repair once the tables are known, so a
// re-write does not surface the Excel-invalid geometry the writer (correctly) rejects.
function dropMergesInsideTables(sheet: Worksheet): void {
  const regions = sheet.tables.map((table) => table.region);
  if (regions.length === 0) return;
  for (const range of [...sheet.merges]) {
    const {top, left, bottom, right} = decodeRange(range);
    if (top === undefined || left === undefined || bottom === undefined || right === undefined)
      continue;
    const overlaps = regions.some(
      (region) =>
        left <= region.right &&
        right >= region.left &&
        top <= region.bottom &&
        bottom >= region.top,
    );
    if (overlaps) sheet.unmergeCells(range);
  }
}

// One `<sheet>` entry from `xl/workbook.xml`: its display name, the rel id linking to the sheet part,
// and its visibility state (absent for a normal, visible sheet).
export interface SheetEntry {
  readonly name: string;
  readonly relId: string;
  readonly state?: WorksheetState['state'];
}

export function parseWorkbookSheets(xml: string): SheetEntry[] {
  const sheets: SheetEntry[] = [];
  for (const {attrs} of openElements(xml, 'sheet')) {
    const entry: {name: string; relId: string; state?: WorksheetState['state']} = {
      name: attrs.name ?? '',
      relId: attrs['r:id'] ?? '',
    };
    if (attrs.state === 'hidden' || attrs.state === 'veryHidden') entry.state = attrs.state;
    sheets.push(entry);
  }
  return sheets;
}

// Read the workbook's structure/window protection (`<workbookProtection>`). The three lock flags are
// decoded as booleans (an absent or "0" attribute stays unlocked), and only the whitelisted
// password/agile-hash attributes are preserved verbatim — a hostile or unknown attribute is dropped
// rather than echoed back on write. Returns undefined when the workbook declares no protection.
export function parseWorkbookProtection(xml: string): WorkbookProtection | undefined {
  let result: WorkbookProtection | undefined;
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) !== 'workbookProtection') return;
      const protection: {
        lockStructure?: boolean;
        lockWindows?: boolean;
        lockRevision?: boolean;
        credentials?: Record<string, string>;
      } = {};
      if (boolStrict(attrs.lockStructure)) protection.lockStructure = true;
      if (boolStrict(attrs.lockWindows)) protection.lockWindows = true;
      if (boolStrict(attrs.lockRevision)) protection.lockRevision = true;
      const credentials: Record<string, string> = {};
      for (const key of WORKBOOK_PROTECTION_CREDENTIAL_ATTRS) {
        const value = attrs[key];
        if (value !== undefined) credentials[key] = value;
      }
      if (Object.keys(credentials).length > 0) protection.credentials = credentials;
      result = protection;
    },
  });
  return result;
}

// Reconstruct the workbook's defined names. Each `<definedName>` carries its name (and optional
// comment/hidden flag) as attributes and its refersTo formula as text content; a `localSheetId`
// maps back through the sheet order to the scope sheet's name. A name whose localSheetId is out of
// range (a foreign file referencing a sheet we did not load) is left global rather than dropped.
function parseWorkbookDefinedNames(xml: string, sheetOrder: readonly string[]): DefinedName[] {
  const names: DefinedName[] = [];
  let capture = false;
  let refersTo = '';
  let pending: {name: string; scope?: string; comment?: string; hidden?: boolean} | undefined;
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) !== 'definedName' || attrs.name === undefined) return;
      // `_xlnm._FilterDatabase` is the built-in Excel derives from a sheet's autofilter, not a
      // user-defined name: it is reconstructed from the sheet's `<autoFilter>` element, so skip it
      // here to keep it off `Workbook.definedNames` and out of a duplicating round-trip.
      if (attrs.name === '_xlnm._FilterDatabase') return;
      capture = true;
      refersTo = '';
      const scopeIndex = attrs.localSheetId === undefined ? -1 : Number(attrs.localSheetId);
      const scope = sheetOrder[scopeIndex];
      pending = {name: attrs.name};
      if (scope !== undefined) pending.scope = scope;
      if (attrs.comment !== undefined) pending.comment = attrs.comment;
      if (boolStrict(attrs.hidden)) pending.hidden = true;
    },
    onText(chunk) {
      if (capture) refersTo += chunk;
    },
    onClose(name) {
      if (localName(name) !== 'definedName' || pending === undefined) return;
      // Strip the `_xlfn.`/`_xlpm.` prefixes back to the readable name, the same normalisation the
      // reader applies to a cell formula, so the model never holds the on-disk mangling.
      names.push({...pending, refersTo: unmangleFunctions(refersTo)});
      capture = false;
      pending = undefined;
    },
  });
  return names;
}

// Shared strings resolve `t="s"` cells. Each `<si>` is one entry: a plain `<si><t>…</t>` decodes to a
// string, while a rich `<si><r><rPr>…</rPr><t>…</t></r>…` decodes to a {@link RichTextValue} whose runs
// carry their per-run fonts — so rich text Excel pooled reads back formatted, not flattened to text.
// The run structure inside an `<si>` is identical to an inline string's `<is>`, so it is parsed the
// same way (see the inline-run accumulation in `parseWorksheet`).
export function parseSharedStrings(xml: string): SharedString[] {
  if (xml === '') return [];
  const strings: SharedString[] = [];
  // Per-`<si>` accumulation: `plain` gathers a bare `<t>`; `runs` gathers `<r>` runs. An `<si>` is
  // rich the moment it holds one `<r>`, at which point its runs — not `plain` — become the entry.
  let plain = '';
  const runs = new RunAccumulator();
  let isRich = false;
  let capture = false;
  let text = '';
  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      switch (local) {
        case 'si':
          plain = '';
          runs.reset();
          isRich = false;
          break;
        case 'r':
          isRich = true;
          runs.beginRun();
          break;
        case 'rPr':
          runs.beginProperties();
          break;
        case 't':
          capture = true;
          text = '';
          break;
        default:
          runs.applyProperty(local, attrs);
          break;
      }
    },
    onText(chunk) {
      if (capture) text += chunk;
    },
    onClose(name) {
      const local = localName(name);
      switch (local) {
        case 't':
          // A `<t>` inside a run is that run's text; a bare `<t>` directly in the `<si>` is plain.
          if (!runs.appendText(text)) plain += text;
          capture = false;
          break;
        case 'r':
          runs.endRun();
          break;
        case 'si':
          strings.push(isRich ? {richText: runs.runs} : plain);
          break;
      }
    },
  });
  return strings;
}

// Core document properties live in docProps/core.xml under mixed namespaces
// (dc:creator, cp:lastModifiedBy, dcterms:created/modified); local names disambiguate.
const CORE_PROPERTY_LOCAL_NAMES = new Set(['creator', 'lastModifiedBy', 'created', 'modified']);

function applyCoreProperties(workbook: Workbook, xml: string): void {
  let capture = '';
  let text = '';
  parseXml(xml, {
    onOpen(name) {
      const local = localName(name);
      capture = CORE_PROPERTY_LOCAL_NAMES.has(local) ? local : '';
      text = '';
    },
    onText(chunk) {
      if (capture !== '') text += chunk;
    },
    onClose(name) {
      if (capture === '' || localName(name) !== capture) return;
      if (capture === 'creator') workbook.properties.creator = text;
      else if (capture === 'lastModifiedBy') workbook.properties.lastModifiedBy = text;
      else {
        const date = new Date(text);
        if (!Number.isNaN(date.getTime())) {
          if (capture === 'created') workbook.properties.created = date;
          else workbook.properties.modified = date;
        }
      }
      capture = '';
    },
  });
}
