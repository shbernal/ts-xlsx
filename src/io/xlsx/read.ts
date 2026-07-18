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
// Untrusted input: inflate is bounded by a running byte counter (`./inflate.ts`) that caps
// actual decompressed output rather than trusting the archive's forgeable size headers, and
// the parser (ADR 0004) never expands entities.

import {strFromU8} from 'fflate';

import {decodeAddress, decodeRange, encodeAddress} from '../../core/address.ts';
import {
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import {
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
  SHEET_PROTECTION_FLAGS,
} from '../../core/protection.ts';
import type {
  Alignment,
  Border,
  BorderStyle,
  Color,
  Fill,
  FillPatternType,
  Font,
  FontVerticalAlignment,
  HorizontalAlignment,
  NamedCellStyle,
  Protection,
  UnderlineStyle,
  VerticalAlignment,
} from '../../core/style.ts';
import {translateFormula, unmangleFunctions} from '../../core/formula.ts';
import type {DataTableFormulaValue, RichTextRun, SharedFormulaValue} from '../../core/value.ts';
import {type DefinedName, Workbook} from '../../core/workbook.ts';
import {
  type WorkbookProtection,
  WORKBOOK_PROTECTION_CREDENTIAL_ATTRS,
} from '../../core/workbook-protection.ts';
import type {
  PageMargins,
  PageSetup,
  PreservedPart,
  PreservedRelationship,
  PreservedWorksheetReference,
  Worksheet,
  WorksheetState,
} from '../../core/worksheet.ts';
import {decodeCellContent, decodeFormulaResult, type SharedString} from './cell-value.ts';
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
import {parseTable} from './tables.ts';
import {localName, parseXml} from './xml-read.ts';

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
const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

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
  partText: (path: string) => string | undefined
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
  partBytes: (path: string) => Uint8Array | undefined
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
  imageIdByMediaPath: Map<string, number>
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
  imageIdByMediaPath: Map<string, number>
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
  sheet: Worksheet
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  const records = parseRelationshipRecords(relsXml);
  const recordById = new Map(records.map(record => [record.id, record]));

  const capture = (
    element: PreservedWorksheetReference['element'],
    relType: string,
    target: string
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
  workbook: Workbook
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
  parseXml(workbookXml, {
    onOpen(name, attrs) {
      if (localName(name) === 'pivotCache' && attrs['r:id'] !== undefined && attrs.cacheId !== undefined) {
        byRelId.set(attrs['r:id'], attrs.cacheId);
      }
    },
    onText() {},
    onClose() {},
  });
  return byRelId;
}

// The `r:id` of the first `<drawing>` / `<legacyDrawingHF>` element in a worksheet, or undefined when
// the sheet declares none. The reference lives in the worksheet XML (not distinguishable by
// relationship Type — a header/footer VML and a comment VML share the `vmlDrawing` type), so the
// specific relationship is found by reading the element's `r:id` here.
function worksheetReferenceRelId(sheetXml: string, element: string): string | undefined {
  let relId: string | undefined;
  parseXml(sheetXml, {
    onOpen(name, attrs) {
      if (relId === undefined && localName(name) === element && attrs['r:id'] !== undefined) {
        relId = attrs['r:id'];
      }
    },
    onText() {},
    onClose() {},
  });
  return relId;
}

// Gather the transitive closure of package parts reachable from an entry part — the part itself, then
// every internal part its relationships target, breadth-first — each with its raw bytes, content type,
// and (internal) relationships. Returns undefined when the entry part is absent (a dangling reference
// preserves nothing). A `visited` set dedupes shared parts and bounds the walk to the (finite,
// inflate-capped) package, so a maliciously self-referential rels graph cannot loop.
function capturePartClosure(
  entryPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  contentTypeOf: (path: string) => string
): readonly PreservedPart[] | undefined {
  const parts: PreservedPart[] = [];
  const visited = new Set<string>();
  const queue: string[] = [entryPath];
  while (queue.length > 0) {
    const path = queue.shift();
    if (path === undefined || visited.has(path)) continue;
    visited.add(path);
    const bytes = partBytes(path);
    if (bytes === undefined) continue;
    const relsXml = partText(relsPathFor(path));
    const rels: PreservedRelationship[] = [];
    if (relsXml !== undefined) {
      for (const rel of parseRelationshipRecords(relsXml)) {
        if (rel.external) continue;
        const targetPath = resolveRelativePart(path, rel.target);
        rels.push({id: rel.id, type: rel.type, targetPath});
        queue.push(targetPath);
      }
    }
    parts.push({path, contentType: contentTypeOf(path), bytes, rels});
  }
  return parts.some(part => part.path === entryPath) ? parts : undefined;
}

// Resolve a package part path to its declared content type the way OPC does: an `<Override>` naming
// the exact part wins, else the `<Default>` registered for the part's extension. An unknown part
// falls back to the generic binary type so re-declaring it never emits an empty content type.
function contentTypeResolver(contentTypesXml: string): (path: string) => string {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  parseXml(contentTypesXml, {
    onOpen(name, attrs) {
      const local = localName(name);
      if (local === 'Override' && attrs.PartName !== undefined && attrs.ContentType !== undefined) {
        overrides.set(attrs.PartName, attrs.ContentType);
      } else if (local === 'Default' && attrs.Extension !== undefined && attrs.ContentType !== undefined) {
        defaults.set(attrs.Extension.toLowerCase(), attrs.ContentType);
      }
    },
    onText() {},
    onClose() {},
  });
  return (path: string): string =>
    overrides.get(`/${path}`) ?? defaults.get(extensionOf(path).toLowerCase()) ?? 'application/octet-stream';
}

// A relationship as declared, with the fields a preserved-part closure needs: its id, Type URI,
// Target, and whether the target lies outside the package (`TargetMode="External"`). A fuller record
// than {@link parseRelationships}'s id→target map, which the closure walk uses to skip external
// targets and carry each relationship's type through a re-write.
function parseRelationshipRecords(
  xml: string
): Array<{id: string; type: string; target: string; external: boolean}> {
  const records: Array<{id: string; type: string; target: string; external: boolean}> = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      if (
        localName(name) === 'Relationship' &&
        attrs.Id !== undefined &&
        attrs.Type !== undefined &&
        attrs.Target !== undefined
      ) {
        records.push({
          id: attrs.Id,
          type: attrs.Type,
          target: attrs.Target,
          external: attrs.TargetMode === 'External',
        });
      }
    },
    onText() {},
    onClose() {},
  });
  return records;
}

// A sheet's tables live in `xl/tables/table{n}.xml` parts, each reached through a relationship of
// type `.../table` on the sheet's own rels. The writer emits one relationship per table; each part
// is parsed back into the model and re-registered in definition order. A part that fails to parse
// (missing name/ref/columns — Excel corruption) is skipped rather than crashing the whole read.
function readSheetTables(
  sheetPath: string,
  partText: (path: string) => string | undefined,
  sheet: Worksheet
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
  sheet: Worksheet
): void {
  const relsXml = partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  for (const target of relationshipTargetsByType(relsXml, 'pivotTable')) {
    const tablePath = resolveRelativePart(sheetPath, target);
    const tableXml = partText(tablePath);
    if (tableXml === undefined) continue;
    const cacheTarget = relationshipTargetByType(partText(relsPathFor(tablePath)) ?? '', 'pivotCacheDefinition');
    const cacheXml = cacheTarget === undefined ? '' : partText(resolveRelativePart(tablePath, cacheTarget)) ?? '';
    sheet.addLoadedPivotTable(parsePivotTable(tableXml, cacheXml));
  }
}

// Excel forbids a merged range inside a formatted table and repairs such a file on load by dropping
// the merge. A worksheet's merges are read before its tables, so a real file carrying that invalid
// geometry lands in the model intact; this applies the same repair once the tables are known, so a
// re-write does not surface the Excel-invalid geometry the writer (correctly) rejects.
function dropMergesInsideTables(sheet: Worksheet): void {
  const regions = sheet.tables.map(table => table.region);
  if (regions.length === 0) return;
  for (const range of [...sheet.merges]) {
    const {top, left, bottom, right} = decodeRange(range);
    if (top === undefined || left === undefined || bottom === undefined || right === undefined) continue;
    const overlaps = regions.some(
      region => left <= region.right && right >= region.left && top <= region.bottom && bottom >= region.top
    );
    if (overlaps) sheet.unmergeCells(range);
  }
}

// The extension of a part path (`xl/media/image1.png` → `png`), or '' when it carries none.
function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf('.');
  return dot === -1 ? '' : partPath.slice(dot + 1);
}

// The relationships for `dir/name.ext` live at `dir/_rels/name.ext.rels`.
function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const base = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// The Target of the first relationship whose Type ends with `/<suffix>` (local-name match, so a
// namespaced or oddly-cased type still resolves), or undefined when none is declared.
function relationshipTargetByType(xml: string, suffix: string): string | undefined {
  let found: string | undefined;
  parseXml(xml, {
    onOpen(name, attrs) {
      if (
        found === undefined &&
        localName(name) === 'Relationship' &&
        attrs.Type !== undefined &&
        attrs.Target !== undefined &&
        attrs.Type.endsWith(`/${suffix}`)
      ) {
        found = attrs.Target;
      }
    },
    onText() {},
    onClose() {},
  });
  return found;
}

// Every Target whose Type ends with `/<suffix>`, in declaration order — for a part class a sheet may
// reference more than once (a sheet can own several tables), where the singular helper's first-match
// would miss all but one.
function relationshipTargetsByType(xml: string, suffix: string): string[] {
  const targets: string[] = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      if (
        localName(name) === 'Relationship' &&
        attrs.Type !== undefined &&
        attrs.Target !== undefined &&
        attrs.Type.endsWith(`/${suffix}`)
      ) {
        targets.push(attrs.Target);
      }
    },
    onText() {},
    onClose() {},
  });
  return targets;
}

// Resolve a relationship target (relative to the referencing part's directory, or absolute from the
// package root) into a package part path, collapsing `.`/`..` segments.
function resolveRelativePart(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseDir = basePart.slice(0, basePart.lastIndexOf('/') + 1);
  const out: string[] = [];
  for (const segment of `${baseDir}${target}`.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}

// A workbook relationship target is relative to the `xl/` directory (`worksheets/sheet1.xml`)
// or absolute from the package root (`/xl/worksheets/sheet1.xml`); normalise both to a part path.
export function resolveWorkbookPart(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  return `xl/${target.replace(/^\.\//, '')}`;
}

export function parseRelationships(xml: string): Map<string, string> {
  const rels = new Map<string, string>();
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) === 'Relationship' && attrs.Id !== undefined && attrs.Target !== undefined) {
        rels.set(attrs.Id, attrs.Target);
      }
    },
    onText() {},
    onClose() {},
  });
  return rels;
}

export function parseWorkbookSheets(
  xml: string,
): Array<{name: string; relId: string; state?: WorksheetState['state']}> {
  const sheets: Array<{name: string; relId: string; state?: WorksheetState['state']}> = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) === 'sheet') {
        const entry: {name: string; relId: string; state?: WorksheetState['state']} = {
          name: attrs.name ?? '',
          relId: attrs['r:id'] ?? '',
        };
        if (attrs.state === 'hidden' || attrs.state === 'veryHidden') entry.state = attrs.state;
        sheets.push(entry);
      }
    },
    onText() {},
    onClose() {},
  });
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
      if (attrs.lockStructure === '1' || attrs.lockStructure === 'true') protection.lockStructure = true;
      if (attrs.lockWindows === '1' || attrs.lockWindows === 'true') protection.lockWindows = true;
      if (attrs.lockRevision === '1' || attrs.lockRevision === 'true') protection.lockRevision = true;
      const credentials: Record<string, string> = {};
      for (const key of WORKBOOK_PROTECTION_CREDENTIAL_ATTRS) {
        const value = attrs[key];
        if (value !== undefined) credentials[key] = value;
      }
      if (Object.keys(credentials).length > 0) protection.credentials = credentials;
      result = protection;
    },
    onText() {},
    onClose() {},
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
      if (attrs.hidden === '1' || attrs.hidden === 'true') pending.hidden = true;
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
  let runs: RichTextRun[] = [];
  let isRich = false;
  let inRun = false;
  let runFont: FontDraft | null = null;
  let runText = '';
  let capture = false;
  let text = '';
  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      switch (local) {
        case 'si':
          plain = '';
          runs = [];
          isRich = false;
          break;
        case 'r':
          isRich = true;
          inRun = true;
          runFont = null;
          runText = '';
          break;
        case 'rPr':
          if (inRun) runFont = {};
          break;
        case 't':
          capture = true;
          text = '';
          break;
        default:
          // A run's `<rPr>` child (`<b/>`, `<sz>`, `<color>`, `<rFont>`, …) is self-closing, read here
          // on open; it sets one font facet. Nothing else uses the default branch.
          if (runFont !== null) applyFontChild(runFont, local, attrs);
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
          if (inRun) runText += text;
          else plain += text;
          capture = false;
          break;
        case 'r':
          if (inRun) {
            const run: {text: string; font?: Partial<Font>} = {text: runText};
            if (runFont !== null && Object.keys(runFont).length > 0) run.font = runFont;
            runs.push(run);
            inRun = false;
          }
          break;
        case 'si':
          strings.push(isRich ? {richText: runs} : plain);
          break;
      }
    },
  });
  return strings;
}

// The style facets an xf resolves to. Absent facets stay undefined, matching the contract
// that an unset facet is simply not present on the reconstructed cell.
export interface XfStyle {
  readonly fill?: Fill;
  readonly numFmt?: string;
  readonly font?: Partial<Font>;
  readonly border?: Border;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
  readonly quotePrefix?: boolean;
  /** The `xfId` link into the named-style layer (`cellStyleXfs`); absent for the Normal default (0). */
  readonly xfId?: number;
}

// A mutable xf accumulator while an <xf> element streams in: its facet ids resolve on open, but
// the <alignment>/<protection> children (when present) arrive before the element closes, so the
// xf is held here and pushed on close rather than on open.
type XfDraft = {-readonly [K in keyof XfStyle]?: XfStyle[K]};

// A mutable font accumulator while a <font> element's children stream in; frozen into a
// Partial<Font> on close.
type FontDraft = {-readonly [K in keyof Font]?: Font[K]};

// A mutable border accumulator while a <border> element's edges stream in; frozen into a
// Border on close. The five edges match Border's; a bare styleless edge is simply never set.
type BorderDraft = {-readonly [K in keyof Border]?: Border[K]};

// The four sides plus the diagonal — the edge elements a <border> can hold, in the order the
// schema lists them. Membership drives edge parsing without a per-name branch.
type BorderEdgeName = 'left' | 'right' | 'top' | 'bottom' | 'diagonal';
const BORDER_EDGES = new Set<string>(['left', 'right', 'top', 'bottom', 'diagonal']);

// ECMA-376 reserves numFmt ids below 164 for formats every consumer knows implicitly, so a
// foreign file may name one with no <numFmt> entry. This maps the standard ids to their
// codes; id 0 (General) and any unknown id resolve to no format. The writer never emits
// these — it always defines a custom id — but reading them keeps foreign files faithful.
const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map([
  [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'],
  [9, '0%'], [10, '0.00%'], [11, '0.00E+00'], [12, '# ?/?'], [13, '# ??/??'],
  [14, 'mm-dd-yy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'], [38, '#,##0 ;[Red](#,##0)'], [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0'], [48, '##0.0E+0'], [49, '@'],
  // Ids 27..36 and 50..58 are reserved for locale-specific built-in East Asian date/time formats;
  // a file authored in a CJK locale styles date cells with them and, being built-ins, emits no
  // <numFmt>. The exact code is locale-defined — these are the representative Excel forms — but what
  // matters for reading is that each resolves to a non-empty date/time code so the serial reads as a
  // date rather than a bare number.
  [27, '[$-404]e/m/d'], [28, '[$-404]e"年"m"月"d"日"'], [29, '[$-404]e"年"m"月"d"日"'], [30, '[$-404]m/d/yy'],
  [31, '[$-404]yyyy"年"m"月"d"日"'], [32, '[$-404]h"時"mm"分"'], [33, '[$-404]h"時"mm"分"ss"秒"'],
  [34, '上午/下午h"時"mm"分"'], [35, '上午/下午h"時"mm"分"ss"秒"'], [36, '[$-404]e/m/d'],
  [50, '[$-404]e/m/d'], [51, '[$-404]e"年"m"月"d"日"'], [52, '[$-404]yyyy"年"m"月"'], [53, '[$-404]m"月"d"日"'],
  [54, '[$-404]e"年"m"月"d"日"'], [55, '上午/下午h"時"mm"分"'], [56, '上午/下午h"時"mm"分"ss"秒"'],
  [57, '[$-404]yyyy"年"m"月"'], [58, '[$-404]m"月"d"日"'],
]);

// styles.xml is a shared table: <numFmts> defines custom format codes by id, <fills> lists
// the fills, and <cellXfs> lists the cell formats, each naming a fill and a number format by
// id. We flatten that indirection into one array — cellXfs index → resolved {fill, numFmt} —
// so a cell/row/column style index maps straight to its facets. The schema orders <numFmts>
// and <fills> before <cellXfs>, so both lookups are complete before an xf references them.
/** The parsed style table: the cell formats a cell/row/column `s` indexes, plus the named cell-style
 * layer a cell's `xfId` links into. Each cellXfs entry is already merged with its named style, so a
 * cell reading its `s` sees the effective facets; the `xfId` link is carried through for re-write. */
export interface StyleTable {
  readonly cellXfs: ReadonlyArray<XfStyle>;
  readonly namedStyles: ReadonlyArray<NamedCellStyle>;
}

export function parseStyleTable(xml: string): StyleTable {
  if (xml === '') return {cellXfs: [], namedStyles: []};
  const fills: Array<Fill | undefined> = [];
  const fonts: Array<Partial<Font> | undefined> = [];
  const borders: Array<Border | undefined> = [];
  const numFmtCodes = new Map<number, string>();
  const xfStyles: XfStyle[] = [];
  // The named-style layer: <cellStyleXfs> holds the base formats a cell's xfId links to; <cellStyles>
  // labels them by name/builtinId. Parsed in parallel with cellXfs, then zipped and merged below.
  const namedXfs: XfStyle[] = [];
  const cellStyleNames: {xfId: number; name?: string; builtinId?: number}[] = [];
  let inFills = false;
  let inFonts = false;
  let inCellXfs = false;
  let inCellStyleXfs = false;
  let pattern = '';
  let fgColor: Color | undefined;
  let bgColor: Color | undefined;
  let fontDraft: FontDraft | null = null;
  let borderDraft: BorderDraft | null = null;
  // Which edge of the current border a nested <color> belongs to; null between edges.
  let currentEdge: BorderEdgeName | null = null;
  // The <cellXfs>/<cellStyleXfs> xf being read; held from open to close so its <alignment>/
  // <protection> children can attach before the xf is committed. null outside an <xf>. `pendingXfTarget`
  // records which table the held xf belongs to.
  let pendingXf: XfDraft | null = null;
  let pendingXfTarget: 'cell' | 'named' | null = null;

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      switch (localName(name)) {
        case 'numFmt': {
          // <numFmts> entries are self-closing, so they are read here on open. A code with
          // no id, or the General id 0, contributes nothing.
          const id = Number(attrs.numFmtId);
          if (Number.isInteger(id) && id > 0 && attrs.formatCode !== undefined) {
            numFmtCodes.set(id, attrs.formatCode);
          }
          break;
        }
        case 'fills':
          inFills = true;
          break;
        case 'fonts':
          inFonts = true;
          break;
        case 'font':
          if (inFonts) {
            fontDraft = {};
            // A bare <font/> would be a font that overrides nothing; keep its id slot.
            if (selfClosing) {
              fonts.push(undefined);
              fontDraft = null;
            }
          }
          break;
        case 'borders':
          break;
        case 'border':
          borderDraft = {};
          currentEdge = null;
          if (attrs.diagonalUp === '1' || attrs.diagonalUp === 'true') borderDraft.diagonalUp = true;
          if (attrs.diagonalDown === '1' || attrs.diagonalDown === 'true') borderDraft.diagonalDown = true;
          // A bare <border/> holds no edges; keep its id slot as the empty border.
          if (selfClosing) {
            borders.push(borderToStyle(borderDraft));
            borderDraft = null;
          }
          break;
        case 'cellXfs':
          inCellXfs = true;
          break;
        case 'cellStyleXfs':
          inCellStyleXfs = true;
          break;
        case 'cellStyle': {
          // A <cellStyle> (inside <cellStyles>) names a cellStyleXfs entry by xfId; it is self-closing,
          // so it is read here on open.
          const xfId = Number(attrs.xfId);
          if (Number.isInteger(xfId)) {
            const entry: {xfId: number; name?: string; builtinId?: number} = {xfId};
            if (attrs.name !== undefined) entry.name = attrs.name;
            if (attrs.builtinId !== undefined) {
              const builtinId = Number(attrs.builtinId);
              if (Number.isInteger(builtinId)) entry.builtinId = builtinId;
            }
            cellStyleNames.push(entry);
          }
          break;
        }
        case 'patternFill':
          if (inFills) {
            pattern = attrs.patternType ?? 'none';
            fgColor = undefined;
            bgColor = undefined;
            // A colourless fill (none/gray125) is a self-closing element, which the SAX
            // parser reports with no matching close — push it here so its id slot is kept.
            if (selfClosing) fills.push(toFill(pattern, undefined, undefined));
          }
          break;
        case 'fgColor':
          if (inFills) fgColor = parseColor(attrs);
          break;
        case 'bgColor':
          if (inFills) bgColor = parseColor(attrs);
          break;
        default: {
          const local = localName(name);
          // A <font>'s children are self-closing, so they are read here on open.
          if (fontDraft !== null) {
            applyFontChild(fontDraft, local, attrs);
          } else if (borderDraft !== null) {
            // A border's edges and their <color> children are all read on open (each is
            // self-closing bar a coloured edge, whose colour child is itself self-closing).
            if (BORDER_EDGES.has(local)) {
              currentEdge = attrs.style !== undefined ? (local as BorderEdgeName) : null;
              if (currentEdge !== null) borderDraft[currentEdge] = {style: attrs.style as BorderStyle};
              if (selfClosing) currentEdge = null;
            } else if (local === 'color' && currentEdge !== null) {
              const edge = borderDraft[currentEdge];
              if (edge !== undefined) borderDraft[currentEdge] = {style: edge.style, color: parseColor(attrs)};
            }
          } else if (pendingXf !== null && local === 'alignment') {
            // An xf's <alignment> child arrives before the xf closes; attach it to the pending xf.
            const alignment = parseAlignment(attrs);
            if (alignment !== undefined) pendingXf.alignment = alignment;
          } else if (pendingXf !== null && local === 'protection') {
            // An xf's <protection> child likewise arrives before the xf closes.
            const protection = parseProtection(attrs);
            if (protection !== undefined) pendingXf.protection = protection;
          } else if ((inCellXfs || inCellStyleXfs) && local === 'xf') {
            // Both <cellXfs> and <cellStyleXfs> hold <xf> with identical structure; they differ only
            // in which table the result lands in and whether an xfId link is meaningful (only cellXfs
            // entries link to a named style).
            const fillId = Number(attrs.fillId);
            const fill = Number.isInteger(fillId) ? fills[fillId] : undefined;
            const fontId = Number(attrs.fontId);
            // Font id 0 is the workbook default font (a real Calibri-11-style face), not an
            // absence — unlike border id 0, which is a genuinely empty border. So an xf naming
            // font 0 resolves to that default face, giving every cell a concrete font to render.
            const font = Number.isInteger(fontId) ? fonts[fontId] : undefined;
            const borderId = Number(attrs.borderId);
            // Border id 0 is the empty default; only a custom border (id > 0) is an explicit one.
            const border = Number.isInteger(borderId) && borderId > 0 ? borders[borderId] : undefined;
            const numFmt = resolveNumFmt(attrs.numFmtId, numFmtCodes);
            const draft: XfDraft = {};
            if (fill) draft.fill = fill;
            if (numFmt !== undefined) draft.numFmt = numFmt;
            if (font) draft.font = font;
            if (border) draft.border = border;
            // The quote-prefix flag is an attribute on the xf itself (no shared sub-table); carry it
            // only when set so an ordinary cell does not gain a spurious `quotePrefix: false`.
            if (attrs.quotePrefix === '1' || attrs.quotePrefix === 'true') draft.quotePrefix = true;
            // A cellXfs entry's xfId links it to a named style; capture it only when it points beyond
            // the Normal default (0), so an ordinary cell carries no spurious named-style link.
            if (inCellXfs && attrs.xfId !== undefined) {
              const xfId = Number(attrs.xfId);
              if (Number.isInteger(xfId) && xfId > 0) draft.xfId = xfId;
            }
            const target = inCellXfs ? xfStyles : namedXfs;
            // A self-closing <xf/> has no alignment child and commits now; otherwise it is held
            // open until its close so an <alignment> child can attach first.
            if (selfClosing) target.push(draft);
            else {
              pendingXf = draft;
              pendingXfTarget = inCellXfs ? 'cell' : 'named';
            }
          }
          break;
        }
      }
    },
    onText() {},
    onClose(name) {
      switch (localName(name)) {
        case 'fills':
          inFills = false;
          break;
        case 'fonts':
          inFonts = false;
          break;
        case 'font':
          if (fontDraft !== null) {
            fonts.push(Object.keys(fontDraft).length > 0 ? fontDraft : undefined);
            fontDraft = null;
          }
          break;
        case 'border':
          if (borderDraft !== null) {
            borders.push(borderToStyle(borderDraft));
            borderDraft = null;
            currentEdge = null;
          }
          break;
        case 'cellXfs':
          inCellXfs = false;
          break;
        case 'cellStyleXfs':
          inCellStyleXfs = false;
          break;
        case 'xf':
          // A held (non-self-closing) xf commits here, with any alignment child attached, into
          // whichever table it was opened in.
          if (pendingXf !== null) {
            (pendingXfTarget === 'named' ? namedXfs : xfStyles).push(pendingXf);
            pendingXf = null;
            pendingXfTarget = null;
          }
          break;
        case 'patternFill':
          if (inFills) fills.push(toFill(pattern, fgColor, bgColor));
          break;
        default:
          // A coloured edge closes after its <color> child; drop the edge context so a stray
          // later <color> cannot attach to it.
          if (borderDraft !== null && BORDER_EDGES.has(localName(name))) currentEdge = null;
          break;
      }
    },
  });

  // Layer each cellXfs entry over the named style its xfId links to: a facet the direct format sets
  // wins; one it leaves unset falls through to the named style. The xfId is carried through so the
  // link survives a re-write. A draft only holds keys for facets it actually set, so the spread merge
  // takes the named base and lets the direct entry override exactly what it names.
  const cellXfs: XfStyle[] = xfStyles.map(xf => {
    if (xf.xfId === undefined) return xf;
    const named = namedXfs[xf.xfId];
    return named === undefined ? xf : {...named, ...xf};
  });

  // Zip the resolved cellStyleXfs facets with their cellStyles name/builtinId into the model's named
  // styles, index for index (a cellStyle's xfId is its cellStyleXfs index).
  const namedStyles: NamedCellStyle[] = namedXfs.map((xf, index) => {
    const label = cellStyleNames.find(entry => entry.xfId === index);
    const style: {-readonly [K in keyof NamedCellStyle]?: NamedCellStyle[K]} = {};
    if (xf.fill !== undefined) style.fill = xf.fill;
    if (xf.numFmt !== undefined) style.numFmt = xf.numFmt;
    if (xf.font !== undefined) style.font = xf.font;
    if (xf.border !== undefined) style.border = xf.border;
    if (xf.alignment !== undefined) style.alignment = xf.alignment;
    if (xf.protection !== undefined) style.protection = xf.protection;
    if (label?.name !== undefined) style.name = label.name;
    if (label?.builtinId !== undefined) style.builtinId = label.builtinId;
    return style;
  });

  return {cellXfs, namedStyles};
}

// A <font> child element sets one facet on the draft. Boolean flags honour their `val`: a
// bare tag or val="1"/"true" is on, val="0"/"false" is off (an explicit-false flag is not
// truthy merely because the tag is present). An unrecognised child is ignored.
function applyFontChild(draft: FontDraft, local: string, attrs: {readonly [k: string]: string}): void {
  switch (local) {
    case 'b':
      draft.bold = flagValue(attrs.val);
      break;
    case 'i':
      draft.italic = flagValue(attrs.val);
      break;
    case 'strike':
      draft.strike = flagValue(attrs.val);
      break;
    case 'outline':
      draft.outline = flagValue(attrs.val);
      break;
    case 'u':
      // A bare <u/> is a single underline; a named style (single/double/…) carries through; but
      // val="none" is the explicit ABSENCE of an underline, so it must read back falsy — not the
      // truthy string "none" that a consumer's `if (font.underline)` would mistake for underlined.
      draft.underline = attrs.val === undefined ? true : attrs.val === 'none' ? false : (attrs.val as UnderlineStyle);
      break;
    case 'vertAlign':
      if (attrs.val !== undefined) draft.vertAlign = attrs.val as FontVerticalAlignment;
      break;
    case 'sz': {
      const size = Number(attrs.val);
      if (Number.isFinite(size)) draft.size = size;
      break;
    }
    case 'color':
      draft.color = parseColor(attrs);
      break;
    // `<name>` in a styles `<font>`, `<rFont>` in a rich-text run's `<rPr>` — the same font face.
    case 'name':
    case 'rFont':
      if (attrs.val !== undefined) draft.name = attrs.val;
      break;
    case 'family': {
      const family = Number(attrs.val);
      if (Number.isInteger(family)) draft.family = family;
      break;
    }
    case 'charset': {
      const charset = Number(attrs.val);
      if (Number.isInteger(charset)) draft.charset = charset;
      break;
    }
    case 'scheme':
      if (attrs.val !== undefined) draft.scheme = attrs.val as Font['scheme'];
      break;
    default:
      break;
  }
}

// An OOXML boolean font flag defaults to on when present with no value (`<b/>` is bold); an
// explicit val turns it on ("1"/"true") or off ("0"/"false").
function flagValue(val: string | undefined): boolean {
  return val === undefined || (val !== '0' && val !== 'false');
}

// Fold a filter column's accumulated `<filters>` or `<customFilters>` state into one criteria value,
// or null when it carried nothing filterable. A `<filters>` block with no values and no blank flag,
// or a `<customFilters>` with no predicates, is a no-op that would round-trip as noise, so it drops.
function pendingFilterCriteria(
  values: string[] | null,
  blank: boolean,
  predicates: CustomFilterPredicate[] | null,
  and: boolean
): FilterCriteria | null {
  if (values !== null && (values.length > 0 || blank)) {
    return {kind: 'values', values, blank};
  }
  if (predicates !== null && predicates.length > 0) {
    return {kind: 'custom', and, predicates: predicates.slice(0, 2)};
  }
  return null;
}

// An xf's numFmtId resolves against the custom codes first, then the built-in table; the
// General format (id 0) and any unrecognised id mean the cell carries no explicit format.
function resolveNumFmt(raw: string | undefined, custom: ReadonlyMap<number, string>): string | undefined {
  if (raw === undefined) return undefined;
  const id = Number(raw);
  if (!Number.isInteger(id) || id === 0) return undefined;
  return custom.get(id) ?? BUILTIN_NUMFMTS.get(id);
}

function toFill(pattern: string, fgColor: Color | undefined, bgColor: Color | undefined): Fill | undefined {
  if (pattern === '' || pattern === 'none') return undefined;
  return {
    type: 'pattern',
    pattern: pattern as FillPatternType,
    ...(fgColor ? {fgColor} : {}),
    ...(bgColor ? {bgColor} : {}),
  };
}

// An accumulated border with no styled edge and no diagonal direction is the empty default:
// it carries nothing, so it resolves to undefined rather than an all-empty Border object.
function borderToStyle(draft: BorderDraft): Border | undefined {
  const hasEdge = (['left', 'right', 'top', 'bottom', 'diagonal'] as const).some(
    (edge: BorderEdgeName): boolean => draft[edge] !== undefined
  );
  if (!hasEdge && draft.diagonalUp === undefined && draft.diagonalDown === undefined) return undefined;
  return draft;
}

// Read an <alignment> element's attributes into an Alignment, keeping only facets that differ
// from the default. Boolean flags honour their parsed value — wrapText="0" is off, so it must
// not fabricate a { wrapText: false } alignment — and an element carrying only defaults yields
// undefined rather than an empty alignment object.
function parseAlignment(attrs: {readonly [k: string]: string}): Alignment | undefined {
  const out: {-readonly [K in keyof Alignment]?: Alignment[K]} = {};
  if (attrs.horizontal !== undefined && attrs.horizontal !== 'general') {
    out.horizontal = attrs.horizontal as HorizontalAlignment;
  }
  if (attrs.vertical !== undefined) out.vertical = attrs.vertical as VerticalAlignment;
  if (attrs.textRotation !== undefined) {
    const rotation = Number(attrs.textRotation);
    if (Number.isFinite(rotation) && rotation !== 0) out.textRotation = rotation;
  }
  if (alignmentFlag(attrs.wrapText)) out.wrapText = true;
  if (attrs.indent !== undefined) {
    const indent = Number(attrs.indent);
    if (Number.isInteger(indent) && indent !== 0) out.indent = indent;
  }
  if (alignmentFlag(attrs.shrinkToFit)) out.shrinkToFit = true;
  if (attrs.readingOrder !== undefined) {
    const order = Number(attrs.readingOrder);
    if (Number.isInteger(order) && order !== 0) out.readingOrder = order;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// An alignment boolean attribute is on only when explicitly "1"/"true"; unlike a font flag it
// is never a bare presence (it always carries a value), and its "0" — a truthy JS string — is off.
function alignmentFlag(val: string | undefined): boolean {
  return val === '1' || val === 'true';
}

// Read a <protection> element into a Protection, keeping only facets that differ from the OOXML
// default. `locked` defaults to TRUE, so only an explicit `locked="0"` carries information (an
// unlocked cell) — a default or explicit-true cell must not read back as { locked: true }; `hidden`
// defaults to false, so only `hidden="1"` is carried. An element with only defaults yields undefined.
function parseProtection(attrs: {readonly [k: string]: string}): Protection | undefined {
  const out: {-readonly [K in keyof Protection]?: Protection[K]} = {};
  if (attrs.locked === '0' || attrs.locked === 'false') out.locked = false;
  if (attrs.hidden === '1' || attrs.hidden === 'true') out.hidden = true;
  return Object.keys(out).length > 0 ? out : undefined;
}

// Read a <sheetProtection> element back into a SheetProtection — the deserialization mirror of the
// writer. `sheet="0"` (or "false") means the element records an *un*protected sheet, so nothing is
// restored. Each flag attribute is the INVERSE of the author's allow-flag ("1" forbids, "0" permits),
// and only attributes actually present are carried, so an omitted (default-valued) flag stays absent —
// exactly what the writer emitted. A password credential is preserved verbatim in its agile form
// (algorithm, hash, salt, spin count); there is no plaintext password to recover, so it is not re-hashed.
function parseSheetProtection(attrs: {readonly [k: string]: string}): SheetProtection | undefined {
  if (attrs.sheet === '0' || attrs.sheet === 'false') return undefined;
  const flags: {-readonly [K in keyof SheetProtectionFlags]?: boolean} = {};
  for (const {key} of SHEET_PROTECTION_FLAGS) {
    const raw = attrs[key];
    if (raw !== undefined) flags[key] = !(raw === '1' || raw === 'true');
  }
  const {algorithmName, hashValue, saltValue, spinCount} = attrs;
  if (algorithmName !== undefined && hashValue !== undefined && saltValue !== undefined && spinCount !== undefined) {
    const spin = Number(spinCount);
    if (Number.isFinite(spin)) {
      const credential: SheetProtectionCredential = {algorithmName, hashValue, saltValue, spinCount: spin};
      return {flags, credential};
    }
  }
  return {flags};
}

function parseColor(attrs: {readonly [k: string]: string}): Color {
  const color: {argb?: string; theme?: number; tint?: number; indexed?: number} = {};
  if (attrs.rgb !== undefined) color.argb = attrs.rgb;
  if (attrs.theme !== undefined) {
    const theme = Number(attrs.theme);
    if (Number.isInteger(theme)) color.theme = theme;
  }
  if (attrs.tint !== undefined) {
    const tint = Number(attrs.tint);
    if (Number.isFinite(tint)) color.tint = tint;
  }
  if (attrs.indexed !== undefined) {
    const indexed = Number(attrs.indexed);
    if (Number.isInteger(indexed)) color.indexed = indexed;
  }
  return color;
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

function parseWorksheet(
  xml: string,
  sheet: Worksheet,
  sharedStrings: readonly SharedString[],
  xfStyles: ReadonlyArray<XfStyle>
): void {
  let cellRef = '';
  let cellType = '';
  let cellStyle = -1;
  let cellCol = -1;
  let cellRow = -1;
  let formula = '';
  // Shared-formula bookkeeping. A master `<f t="shared" ref si>TEXT</f>` seeds the group; every clone
  // `<f t="shared" si/>` in the sheet references it by `si` and carries no text of its own. Masters
  // always precede their clones (Excel keeps the master top-left), so a clone resolves against a map
  // filled as the sheet streams: the master's formula translated to the clone's position.
  const sharedMasters = new Map<number, {formula: string; col: number; row: number}>();
  let formulaShared = false;
  let formulaSi = -1;
  let sharedClone = false;
  // The declaration attributes of a `<f t="dataTable">`, held from the `<f>` open until the cell
  // finalises. Null for every other cell.
  let formulaDataTable: {ref: string; dt2D: string | undefined; dtr: string | undefined; r1: string | undefined; r2: string | undefined} | null =
    null;
  let valueText = '';
  let inlineText = '';
  let hasFormula = false;
  let hasValue = false;
  let inInlineString = false;
  let capture = false;
  let text = '';
  // Rich-text run accumulation while inside an `<is>` built from `<r>` runs: the runs gathered so
  // far, the current run's font draft (null until an `<rPr>` opens) and its text, and whether we are
  // inside an `<r>`. A `<t>` inside a run appends to the run text; a bare `<t>` in the `<is>` appends
  // to `inlineText`, keeping a plain inline string on its existing path.
  let inlineRuns: RichTextRun[] = [];
  let runFont: {-readonly [K in keyof Font]?: Font[K]} | null = null;
  let runText = '';
  let inRun = false;
  // A row with customFormat="1" supplies a default style for its cells that carry no `s`.
  let rowStyle = -1;
  let rowCustomFormat = false;
  // Autofilter accumulation. The sheet `<autoFilter ref>` seeds a draft; each `<filterColumn colId>`
  // opens a column whose criteria (`<filters>` values or `<customFilters>` predicates) stream into
  // the draft until `</filterColumn>`, and `</autoFilter>` commits the whole thing to the sheet.
  let filterRef: string | null = null;
  let filterColumns: FilterColumn[] = [];
  let filterColId = -1;
  let filterValues: string[] | null = null;
  let filterBlank = false;
  let customPredicates: CustomFilterPredicate[] | null = null;
  let customAnd = false;
  // `<brk>` elements appear under both `<rowBreaks>` and `<colBreaks>`; this flag routes only the
  // row-break children onto the model (column breaks are not modelled yet).
  let inRowBreaks = false;
  // A column's `style` is the default for its cells that carry no style of their own; this
  // maps a column index to that style index so a bare cell can inherit it (as Excel does,
  // without stamping every cell). Columns are parsed before any cell references them.
  const columnStyle = new Map<number, number>();

  // Commit the accumulated autofilter to the sheet. Shared by the `</autoFilter>` close and the
  // self-closing `<autoFilter/>` open (a criteria-free filter fires no close). Columns whose colId
  // falls outside the range are dropped here so the strict setter never trips on hostile input.
  const commitAutoFilter = (): void => {
    if (filterRef === null) return;
    try {
      const {left, right} = decodeRange(filterRef);
      const width = left !== undefined && right !== undefined ? right - left + 1 : 0;
      sheet.autoFilter = {ref: filterRef, columns: filterColumns.filter(c => c.colId < width)};
    } catch {
      // unbounded or malformed autofilter range in the source file — ignore it
    }
    filterRef = null;
    filterColumns = [];
  };

  // Commit the cell held in the parser state to the sheet, resolving its style from its own `s`,
  // then its row's (when customFormat), then its column's default — the order Excel applies. Shared
  // by the normal `</c>` close and the self-closing `<c/>` open of a formatted-but-empty cell.
  const finalizeCellFromState = (): void => {
    if (cellRef === '') return;
    const styleIndex =
      cellStyle >= 0
        ? cellStyle
        : rowCustomFormat && rowStyle >= 0
          ? rowStyle
          : columnStyle.get(cellCol) ?? -1;
    const style = styleIndex >= 0 ? xfStyles[styleIndex] : xfStyles[0];
    // A data-table formula is preserved by declaration, not evaluated: surface its kind, range, input
    // cells, and cached result so a read-modify-write cycle re-emits it rather than dropping it.
    if (formulaDataTable !== null) {
      const value: DataTableFormulaValue = {
        shareType: 'dataTable',
        ref: formulaDataTable.ref,
        ...(flagValue(formulaDataTable.dt2D ?? '0') ? {dataTable2D: true} : {}),
        ...(flagValue(formulaDataTable.dtr ?? '0') ? {dataTableRow: true} : {}),
        ...(formulaDataTable.r1 !== undefined ? {r1: formulaDataTable.r1} : {}),
        ...(formulaDataTable.r2 !== undefined ? {r2: formulaDataTable.r2} : {}),
        ...(hasValue ? {result: decodeFormulaResult(cellType, valueText, style?.numFmt)} : {}),
      };
      const dtCell = sheet.getCell(cellRef);
      applyCellStyle(dtCell, style);
      dtCell.value = value;
      return;
    }
    // A shared-formula master seeds the group for its clones before it is finalised as an ordinary
    // formula cell; a clone resolves to the master's formula translated to its own position and is
    // committed here directly, since its value is not a plain `<c>` payload decodeCellContent knows.
    if (hasFormula && formulaShared && formulaSi >= 0) {
      sharedMasters.set(formulaSi, {formula, col: cellCol, row: cellRow});
    } else if (sharedClone && formulaSi >= 0) {
      const master = sharedMasters.get(formulaSi);
      if (master !== undefined) {
        const translated = translateFormula(master.formula, cellCol - master.col, cellRow - master.row);
        const value: SharedFormulaValue = {
          sharedFormula: encodeAddress(master.col, master.row),
          formula: unmangleFunctions(translated),
          // A clone's cached result honours the cell's date format the same way a plain formula's does.
          ...(hasValue ? {result: decodeFormulaResult(cellType, valueText, style?.numFmt)} : {}),
        };
        const cloneCell = sheet.getCell(cellRef);
        applyCellStyle(cloneCell, style);
        cloneCell.value = value;
        return;
      }
    }
    finalizeCell(sheet, cellRef, cellType, hasFormula, formula, hasValue, valueText, inlineText, inlineRuns, sharedStrings, style);
  };

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      const local = localName(name);
      text = '';
      capture = false;
      switch (local) {
        case 'col':
          applyColumn(sheet, attrs, xfStyles, columnStyle);
          break;
        case 'row':
          applyRow(sheet, attrs);
          rowStyle = attrs.s !== undefined ? Number(attrs.s) : -1;
          rowCustomFormat = attrs.customFormat === '1' || attrs.customFormat === 'true';
          break;
        case 'c':
          cellRef = attrs.r ?? '';
          cellType = attrs.t ?? '';
          cellStyle = attrs.s !== undefined ? Number(attrs.s) : -1;
          cellCol = cellRef === '' ? -1 : decodeAddress(cellRef).col ?? -1;
          cellRow = cellRef === '' ? -1 : decodeAddress(cellRef).row ?? -1;
          formula = '';
          valueText = '';
          inlineText = '';
          inlineRuns = [];
          inRun = false;
          runFont = null;
          hasFormula = false;
          hasValue = false;
          formulaShared = false;
          formulaSi = -1;
          sharedClone = false;
          formulaDataTable = null;
          // A self-closing `<c r=".." s=".."/>` is a formatted-but-empty cell: it carries a style but
          // no value, so no `close` fires to finalise it. Commit it here from its style alone, else
          // the formatting on a blank cell is silently lost on read.
          if (selfClosing) finalizeCellFromState();
          break;
        case 'is':
          inInlineString = true;
          inlineText = '';
          inlineRuns = [];
          break;
        case 'r':
          // A run inside a rich inline string. Its `<rPr>` (if any) and `<t>` follow; reset the
          // per-run accumulators so an unformatted run inherits nothing from the previous one.
          if (inInlineString) {
            inRun = true;
            runFont = null;
            runText = '';
          }
          break;
        case 'rPr':
          // The run's formatting bundle; its self-closing children stream into the default branch.
          if (inRun) runFont = {};
          break;
        case 'f':
          capture = true;
          formulaShared = attrs.t === 'shared';
          formulaSi = attrs.si !== undefined ? Number(attrs.si) : -1;
          // A self-closing `<f t="shared" si/>` is a clone: it fires no close event and carries no
          // text, so mark it here to resolve against its master when the cell finalises.
          if (selfClosing && formulaShared) sharedClone = true;
          // A `<f t="dataTable">` carries only declaration attributes; hold them for finalisation so
          // the data-table kind is preserved rather than read as an empty formula.
          if (attrs.t === 'dataTable' && attrs.ref !== undefined) {
            formulaDataTable = {ref: attrs.ref, dt2D: attrs.dt2D, dtr: attrs.dtr, r1: attrs.r1, r2: attrs.r2};
          }
          break;
        case 'v':
        case 't':
          capture = true;
          break;
        case 'oddHeader':
        case 'oddFooter':
        case 'evenHeader':
        case 'evenFooter':
        case 'firstHeader':
        case 'firstFooter':
          // A `<headerFooter>` child carries its header/footer definition as text (the `&`-prefixed
          // section/format tokens, e.g. `&C&"Arial"&G`). Capture it verbatim so a round-trip preserves
          // a header image's `&G` picture token and every other formatting directive.
          capture = true;
          break;
        case 'mergeCell':
          // A well-formed file never declares overlapping merges; a corrupt one might. Reject the
          // bad range at the model boundary, but don't let one abort the whole parse — drop it and
          // keep reading the valid geometry.
          if (attrs.ref !== undefined && attrs.ref !== '') {
            try {
              sheet.mergeCells(attrs.ref);
            } catch {
              // overlapping/malformed merge in the source file — skip it
            }
          }
          break;
        case 'tabColor':
          // A self-closing `<sheetPr>` child, so it arrives here rather than as text.
          sheet.tabColor = parseColor(attrs);
          break;
        case 'outlinePr':
          // Another self-closing `<sheetPr>` child; only set flags the source actually carried, so
          // a file without them leaves `outline` empty and a re-write stays byte-clean.
          if (attrs.summaryBelow !== undefined) sheet.outline.summaryBelow = flagValue(attrs.summaryBelow);
          if (attrs.summaryRight !== undefined) sheet.outline.summaryRight = flagValue(attrs.summaryRight);
          break;
        case 'pane':
          // A `<sheetView>` child recording a frozen (or split) pane. Only a frozen pane maps onto
          // the model's view; a source without one leaves `view` empty, so a re-write emits no pane.
          if (attrs.state === 'frozen' || attrs.state === 'frozenSplit') {
            sheet.view.state = 'frozen';
            if (attrs.xSplit !== undefined) sheet.view.xSplit = Number(attrs.xSplit);
            if (attrs.ySplit !== undefined) sheet.view.ySplit = Number(attrs.ySplit);
            if (attrs.topLeftCell !== undefined) sheet.view.topLeftCell = attrs.topLeftCell;
          }
          break;
        case 'pageSetUpPr':
          // The fit-to-page flag, a self-closing `<sheetPr>` child. Recorded only when the source
          // carried the attribute, so a `<pageSetUpPr>` present for other reasons (e.g.
          // `autoPageBreaks`) leaves `pageSetup.fitToPage` unset.
          if (attrs.fitToPage !== undefined) sheet.pageSetup.fitToPage = flagValue(attrs.fitToPage);
          break;
        case 'pageMargins':
          applyMargins(sheet.pageMargins, attrs);
          break;
        case 'pageSetup':
          applyPageSetup(sheet.pageSetup, attrs);
          break;
        case 'rowBreaks':
          inRowBreaks = true;
          break;
        case 'colBreaks':
          // A `<brk>` under column breaks must not land on the row-break model; a self-closing
          // `<colBreaks/>` fires no close, so the flag is simply left false here.
          inRowBreaks = false;
          break;
        case 'brk': {
          // Only manual row breaks are surfaced. `id` is the row the layout splits before; a
          // non-positive or non-integer id is hostile input and dropped rather than trusted.
          if (!inRowBreaks) break;
          const id = Number(attrs.id);
          if (!Number.isInteger(id) || id < 1) break;
          const brk: {id: number; max?: number; man?: boolean} = {id};
          const max = Number(attrs.max);
          if (Number.isInteger(max) && max >= 0) brk.max = max;
          if (attrs.man === '1' || attrs.man === 'true') brk.man = true;
          sheet.rowBreaks.push(brk);
          break;
        }
        case 'sheetProtection': {
          const protection = parseSheetProtection(attrs);
          if (protection !== undefined) sheet.restoreProtection(protection);
          break;
        }
        case 'autoFilter':
          // Seed a draft from the range; its `<filterColumn>` children (if any) stream in below and
          // `</autoFilter>` commits it. A criteria-free filter is self-closing and fires no close, so
          // commit it here. (A table's own `<autoFilter>` also matches, but table sheets route
          // through parseTable, so this only ever sees the sheet-level one.)
          filterRef = attrs.ref !== undefined && attrs.ref !== '' ? attrs.ref : null;
          filterColumns = [];
          if (selfClosing) commitAutoFilter();
          break;
        case 'filterColumn':
          // A criteria block for one column, offset `colId` from the range's left edge. Reset the
          // per-column accumulators; whichever child (`<filters>`/`<customFilters>`) opens fills one.
          filterColId = attrs.colId !== undefined ? Number(attrs.colId) : -1;
          filterValues = null;
          filterBlank = false;
          customPredicates = null;
          customAnd = false;
          break;
        case 'filters':
          filterValues = [];
          filterBlank = flagValue(attrs.blank) && attrs.blank !== undefined;
          break;
        case 'filter':
          if (filterValues !== null && attrs.val !== undefined) filterValues.push(attrs.val);
          break;
        case 'customFilters':
          customPredicates = [];
          customAnd = attrs.and !== undefined && flagValue(attrs.and);
          break;
        case 'customFilter': {
          // The operator attribute defaults to `equal` when absent (per CT_CustomFilter); an operand
          // is likewise optional. An unrecognised operator drops the predicate rather than guessing.
          if (customPredicates === null) break;
          const operator = attrs.operator ?? 'equal';
          if (isCustomFilterOperator(operator)) {
            customPredicates.push({operator, val: attrs.val ?? ''});
          }
          break;
        }
        default:
          // A run's `<rPr>` child (`<b/>`, `<sz>`, `<color>`, `<rFont>`, …) sets one font facet; it
          // is self-closing, so it is read here on open. Nothing else uses the default branch.
          if (runFont !== null) applyFontChild(runFont, local, attrs);
          break;
      }
      if (selfClosing && (local === 'f' || local === 'v')) capture = false;
    },
    onText(chunk) {
      if (capture) text += chunk;
    },
    onClose(name) {
      const local = localName(name);
      switch (local) {
        case 'f':
          formula = text;
          hasFormula = true;
          break;
        case 'v':
          valueText = text;
          hasValue = true;
          break;
        case 't':
          // A `<t>` inside a run is that run's text; a bare `<t>` directly in the `<is>` is a plain
          // inline string. `inRun` must be checked first — a run is also inside the inline string.
          if (inRun) runText += text;
          else if (inInlineString) inlineText += text;
          break;
        case 'r':
          if (inRun) {
            const run: {text: string; font?: Partial<Font>} = {text: runText};
            if (runFont !== null && Object.keys(runFont).length > 0) run.font = runFont;
            inlineRuns.push(run);
            inRun = false;
          }
          break;
        case 'is':
          inInlineString = false;
          break;
        case 'oddHeader':
        case 'oddFooter':
        case 'evenHeader':
        case 'evenFooter':
        case 'firstHeader':
        case 'firstFooter':
          sheet.headerFooter[local] = text;
          break;
        case 'c':
          finalizeCellFromState();
          break;
        case 'row':
          rowStyle = -1;
          rowCustomFormat = false;
          break;
        case 'filterColumn': {
          // Assemble this column's criteria from whichever accumulator filled. A column whose colId
          // is negative, or whose criteria are empty (no values, no blank, no predicates), carries
          // nothing filterable and is dropped so a re-write stays clean — load-repair, not authoring.
          const criteria = pendingFilterCriteria(
            filterValues,
            filterBlank,
            customPredicates,
            customAnd
          );
          if (filterColId >= 0 && criteria !== null) {
            filterColumns.push({colId: filterColId, criteria});
          }
          break;
        }
        case 'autoFilter':
          commitAutoFilter();
          break;
        case 'rowBreaks':
          inRowBreaks = false;
          break;
        default:
          break;
      }
      capture = false;
    },
  });
}

function applyColumn(
  sheet: Worksheet,
  attrs: {readonly [k: string]: string},
  xfStyles: ReadonlyArray<XfStyle>,
  columnStyle: Map<number, number>
): void {
  const min = Number(attrs.min);
  const max = Number(attrs.max);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1) return;
  const width = attrs.width !== undefined ? Number(attrs.width) : undefined;
  const hidden = attrs.hidden === '1' || attrs.hidden === 'true';
  const styleIndex = attrs.style !== undefined ? Number(attrs.style) : -1;
  // The column's style resolves to the same facet bundle a cell's does; mirror all of it onto the
  // column model so `getColumn(i)` reflects the declared default, not just its number format.
  const style = styleIndex >= 0 ? xfStyles[styleIndex] : undefined;
  for (let index = min; index <= max; index++) {
    const properties = sheet.getColumn(index);
    if (width !== undefined && Number.isFinite(width) && attrs.customWidth !== '0') properties.width = width;
    if (hidden) properties.hidden = true;
    if (attrs.outlineLevel !== undefined) {
      const level = Number(attrs.outlineLevel);
      if (Number.isInteger(level) && level > 0) properties.outlineLevel = level;
    }
    if (attrs.collapsed === '1' || attrs.collapsed === 'true') properties.collapsed = true;
    if (style?.numFmt !== undefined) properties.numFmt = style.numFmt;
    if (style?.fill !== undefined) properties.fill = style.fill;
    if (style?.font !== undefined) properties.font = style.font;
    if (style?.border !== undefined) properties.border = style.border;
    if (style?.alignment !== undefined) properties.alignment = style.alignment;
    if (style?.protection !== undefined) properties.protection = style.protection;
    // Record the column's style so a bare cell in it can inherit the full column format on read.
    if (styleIndex >= 0) columnStyle.set(index, styleIndex);
  }
}

function applyRow(sheet: Worksheet, attrs: {readonly [k: string]: string}): void {
  const number = Number(attrs.r);
  if (!Number.isInteger(number) || number < 1) return;
  const properties = sheet.getRow(number);
  if (attrs.ht !== undefined && attrs.customHeight !== '0') {
    const height = Number(attrs.ht);
    if (Number.isFinite(height)) properties.height = height;
  }
  if (attrs.hidden === '1' || attrs.hidden === 'true') properties.hidden = true;
  if (attrs.outlineLevel !== undefined) {
    const level = Number(attrs.outlineLevel);
    if (Number.isInteger(level) && level > 0) properties.outlineLevel = level;
  }
  if (attrs.collapsed === '1' || attrs.collapsed === 'true') properties.collapsed = true;
}

function applyMargins(margins: PageMargins, attrs: {readonly [k: string]: string}): void {
  for (const side of MARGIN_SIDES) {
    const raw = attrs[side];
    if (raw === undefined) continue;
    const value = Number(raw);
    if (Number.isFinite(value)) margins[side] = value;
  }
}

// Read the `<pageSetup>` print-scaling attributes back onto the model, setting only those the
// source carried so a re-write stays byte-clean. Numeric attributes that fail to parse are
// dropped rather than stored as NaN; the enumerated ones are trusted verbatim (an unexpected token
// round-trips harmlessly as an unknown string).
function applyPageSetup(pageSetup: PageSetup, attrs: {readonly [k: string]: string}): void {
  const num = (raw: string | undefined): number | undefined => {
    if (raw === undefined) return undefined;
    const value = Number(raw);
    return Number.isFinite(value) ? value : undefined;
  };
  const paperSize = num(attrs.paperSize);
  if (paperSize !== undefined) pageSetup.paperSize = paperSize;
  const scale = num(attrs.scale);
  if (scale !== undefined) pageSetup.scale = scale;
  const fitToWidth = num(attrs.fitToWidth);
  if (fitToWidth !== undefined) pageSetup.fitToWidth = fitToWidth;
  const fitToHeight = num(attrs.fitToHeight);
  if (fitToHeight !== undefined) pageSetup.fitToHeight = fitToHeight;
  if (attrs.pageOrder === 'downThenOver' || attrs.pageOrder === 'overThenDown') {
    pageSetup.pageOrder = attrs.pageOrder;
  }
  if (attrs.orientation === 'portrait' || attrs.orientation === 'landscape') {
    pageSetup.orientation = attrs.orientation;
  }
}

function finalizeCell(
  sheet: Worksheet,
  ref: string,
  type: string,
  hasFormula: boolean,
  formula: string,
  hasValue: boolean,
  valueText: string,
  inlineText: string,
  richTextRuns: readonly RichTextRun[],
  sharedStrings: readonly SharedString[],
  style: XfStyle | undefined
): void {
  const {col, row} = decodeAddress(ref);
  if (col === undefined || row === undefined) return;
  const cell = sheet.getCell(ref);
  applyCellStyle(cell, style);

  cell.value = decodeCellContent(
    {type, hasFormula, formula, hasValue, valueText, inlineText, richTextRuns},
    sharedStrings,
    style?.numFmt
  );
}

// Applies a resolved xf's non-value facets to a cell. Shared by the ordinary cell path and the
// shared-formula clone path, so a styled clone (fill/font/border/alignment/protection) keeps its
// look on read rather than surviving as value-only.
function applyCellStyle(cell: Cell, style: XfStyle | undefined): void {
  if (style?.fill !== undefined) cell.fill = style.fill;
  if (style?.numFmt !== undefined) cell.numFmt = style.numFmt;
  if (style?.font !== undefined) cell.font = style.font;
  if (style?.border !== undefined) cell.border = style.border;
  if (style?.alignment !== undefined) cell.alignment = style.alignment;
  if (style?.protection !== undefined) cell.protection = style.protection;
  if (style?.quotePrefix !== undefined) cell.quotePrefix = style.quotePrefix;
  if (style?.xfId !== undefined) cell.namedStyleId = style.xfId;
}
