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

import {decodeRange} from '../../core/address.ts';
import type {CommentThread} from '../../core/comment-thread.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import type {PreservedWorksheetReference} from '../../core/preserved.ts';
import {type DefinedName, Workbook, type WorkbookView} from '../../core/workbook.ts';
import {
  WORKBOOK_PROTECTION_CREDENTIAL_ATTRS,
  type WorkbookProtection,
  type WorkbookProtectionCredentialAttr,
} from '../../core/workbook-protection.ts';
import type {Worksheet, WorksheetState} from '../../core/worksheet.ts';
import {applyNotes, type ParsedComment, parseComments} from './comments.ts';
import {parseConditionalFormattings, parseDxfs} from './conditional-formatting.ts';
import {
  applyDataValidations,
  parseDataValidations,
  parseExtendedDataValidations,
} from './data-validation.ts';
import {applyHyperlinks, parseSheetHyperlinks} from './hyperlinks.ts';
import {drawingHasUnmodeledContent, parseDrawing} from './images.ts';
import {extensionOf} from './part-paths.ts';
import {parsePivotTable} from './pivot-read.ts';
import {
  capturePartClosure,
  contentTypeResolver,
  type PackageAccessors,
  packageAccessors,
  parseRelationshipRecords,
  parseRelationships,
  relationshipTargetByType,
  relationshipTargetsByType,
  relsPathFor,
  resolveRelativePart,
  resolveWorkbookPart,
  sheetRelTarget,
} from './read-opc.ts';
import {parseStyleTable} from './read-styles.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {parseSharedStrings} from './shared-strings-read.ts';
import {inflateXlsxPackage, unsupportedWorkbookPart} from './sniff-format.ts';
import {parseIndexedColors, parseMruColors, parseTableStyles} from './styles.ts';
import {parseTable} from './tables.ts';
import {buildCommentThreads, parsePersons, parseThreadedComments} from './threaded-comments.ts';
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
 * @throws {UnsupportedFormatError} if the input is not a readable `.xlsx` package — a legacy `.xls`
 *   (`.format === 'xls'`), a binary `.xlsb` (`'xlsb'`), or an unrecognised/non-ZIP blob (`'unknown'`).
 * @throws {Error} if the archive exceeds the inflate bound (a probable zip bomb).
 */
export function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook {
  const cap = options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  const pkg = packageAccessors(inflateXlsxPackage(data, cap));
  const {partText} = pkg;

  const workbookXml = partText('xl/workbook.xml');
  if (workbookXml === undefined) throw unsupportedWorkbookPart(partText);

  // A part's content type is needed to faithfully re-declare any part preserved verbatim for
  // round-tripping (a vector-shape drawing, a header/footer image and its VML). Resolve it the way
  // OPC does: an explicit `<Override>` for the exact part, else the `<Default>` for its extension.
  const contentTypeOf = contentTypeResolver(partText('[Content_Types].xml') ?? '');

  const workbookRelsXml = partText('xl/_rels/workbook.xml.rels') ?? '';
  const rels = parseRelationships(workbookRelsXml);
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
  // Preserve the author's "Recent Colors" swatches, which the model never reads but re-writing would
  // otherwise discard.
  workbook.restoreMruColors(parseMruColors(stylesXml));
  // Preserve the custom table-style definitions so a table referencing one by name still resolves to
  // a real definition after a re-write instead of rendering unstyled.
  workbook.restoreTableStyles(parseTableStyles(stylesXml));
  // Preserve the theme part so a branded colour/font scheme is not overwritten by the default theme
  // the writer emits for a workbook that has none.
  readWorkbookTheme(workbookRelsXml, pkg, contentTypeOf, workbook);
  // Preserve the named cell-style layer only when a file declares one beyond the Normal default, so an
  // ordinary workbook keeps an empty named-style table and emits just the default on write.
  if (namedStyles.length > 1) workbook.restoreNamedStyles(namedStyles);
  const core = partText('docProps/core.xml');
  if (core !== undefined) applyCoreProperties(workbook, core);
  workbook.protection = parseWorkbookProtection(workbookXml);
  applyWorkbookView(workbook.view, workbookXml);
  // The threaded-comment author registry is workbook-level, and every conversation on every sheet
  // resolves its authors and @mentions through it — so it is restored before the sheet loop that reads
  // those conversations, not alongside the other workbook-level parts below.
  readWorkbookPersons(workbookRelsXml, pkg, workbook);

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
      // Threads before notes: a threaded cell's comments-part entry is the thread's legacy fallback, not
      // a note, and `applyNotes` reads the sheet's restored threads to tell the two apart.
      const threads = readSheetCommentThreads(path, pkg, workbook);
      if (threads.length > 0) sheet.restoreCommentThreads(threads);
      const comments = readSheetComments(path, pkg);
      if (comments !== undefined) applyNotes(sheet, comments);
      readSheetImages(path, pkg, workbook, sheet, imageIdByMediaPath);
      readSheetBackground(path, pkg, workbook, sheet, imageIdByMediaPath);
      if (sheetXml !== undefined) {
        readSheetPreservedReferences(path, sheetXml, pkg, contentTypeOf, sheet);
      }
      readSheetTables(path, pkg, sheet);
      readSheetPivotTables(path, pkg, sheet);
      const printerSettings = readSheetPrinterSettings(path, pkg);
      if (printerSettings !== undefined) sheet.pageSetup.printerSettings = printerSettings;
    }
  }

  readWorkbookPreservedReferences(workbookXml, pkg, contentTypeOf, workbook);
  readRootPreservedReferences(pkg, contentTypeOf, workbook);

  // Defined names follow the sheets: a scoped name's `localSheetId` indexes the sheet order, which
  // is why the names are read only once every sheet is registered.
  for (const name of parseWorkbookDefinedNames(workbookXml, sheetOrder)) {
    workbook.defineName(name);
  }
  return workbook;
}

// A sheet's comments live in a comments part reached through the sheet's own relationships: the sheet
// declares a relationship of type `.../comments` whose target resolves (relative to the sheet's
// directory) to the comments part. A sheet with no rels part or no such relationship simply has none.
function readSheetComments(
  sheetPath: string,
  pkg: PackageAccessors,
): Map<string, ParsedComment> | undefined {
  const commentsPath = sheetRelTarget(sheetPath, pkg.partText, 'comments');
  if (commentsPath === undefined) return undefined;
  const commentsXml = pkg.partText(commentsPath);
  if (commentsXml === undefined) return undefined;
  return parseComments(commentsXml);
}

// The workbook's threaded-comment identity registry: a relationship of type `.../person` names
// `xl/persons/person.xml`, whose entries every message's `personId` and every mention's
// `mentionpersonId` resolve through. A workbook with no threaded comments declares no such
// relationship and keeps an empty registry.
function readWorkbookPersons(
  workbookRelsXml: string,
  pkg: PackageAccessors,
  workbook: Workbook,
): void {
  const target = relationshipTargetByType(workbookRelsXml, 'person');
  const xml = target === undefined ? undefined : pkg.partText(resolveWorkbookPart(target));
  if (xml !== undefined) workbook.restorePersons(parsePersons(xml));
}

// The workbook's theme part: the `<clrScheme>`/`<fontScheme>`/`<fmtScheme>` every `theme="n"` colour
// and every `scheme="major|minor"` font in the package resolves against. It is reached through the
// workbook's `.../theme` relationship rather than assumed at `xl/theme/theme1.xml`, because the target
// is rel-relative and a foreign package is free to name the part anything.
//
// Captured with its transitive part closure, not as a lone string: a theme can carry its own
// relationships (a picture used as a themed fill, wired by an `r:embed` into the theme's rels part),
// and re-emitting the theme body without them would leave that reference dangling — which Excel
// reports as a package needing repair. A package that declares no theme leaves the workbook on the
// library's default, which is also what a dangling relationship target degrades to.
function readWorkbookTheme(
  workbookRelsXml: string,
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const target = relationshipTargetByType(workbookRelsXml, 'theme');
  if (target === undefined) return;
  const entryPath = resolveWorkbookPart(target);
  const parts = capturePartClosure(entryPath, pkg.partText, pkg.partBytes, contentTypeOf);
  if (parts !== undefined) workbook.restoreThemePart({entryPath, parts});
}

// A sheet's threaded conversations live in a `xl/threadedComments/threadedComment{n}.xml` part reached
// through a relationship of type `.../threadedComment` on the sheet's own rels — the same discovery
// shape as the notes part above, and deliberately separate from it: a thread and a legacy note are
// different features that happen to share a sheet. The messages are grouped into threads and their
// authors resolved against the workbook registry, so each thread lands self-contained.
//
// What lands here IS what a re-write emits: the thread part is re-serialised from these threads, and so is
// the legacy fallback `<comment>` that binds each cell to its conversation (see `comments.ts`). Anything
// this reader drops is therefore dropped from the file — which is why a message too damaged to place is
// still kept wherever it can be, and why the anchor is canonicalised here rather than trusted downstream.
function readSheetCommentThreads(
  sheetPath: string,
  pkg: PackageAccessors,
  workbook: Workbook,
): CommentThread[] {
  const path = sheetRelTarget(sheetPath, pkg.partText, 'threadedComment');
  const xml = path === undefined ? undefined : pkg.partText(path);
  if (xml === undefined) return [];
  return buildCommentThreads(parseThreadedComments(xml), (id) => workbook.getPerson(id));
}

// A sheet's printer-settings blob is an opaque binary part linked from `<pageSetup r:id>`: the sheet
// declares a relationship of type `.../printerSettings` whose target resolves to a `.bin` part. We
// keep the raw bytes verbatim — the DEVMODE inside is platform-specific and the model never
// interprets it, only round-trips it so re-writing the file preserves the user's print configuration.
// A sheet with no rels part or no such relationship simply has none.
function readSheetPrinterSettings(
  sheetPath: string,
  pkg: PackageAccessors,
): Uint8Array | undefined {
  const path = sheetRelTarget(sheetPath, pkg.partText, 'printerSettings');
  return path === undefined ? undefined : pkg.partBytes(path);
}

// A sheet's anchored images live in a drawing part reached through the sheet's own relationships: a
// relationship of type `.../drawing` names the drawing part, whose own relationships map each
// picture's embed id to a media part under `xl/media/`. Each anchor becomes a workbook image (deduped
// by media path) placed back on the sheet at its two-cell anchor.
function readSheetImages(
  sheetPath: string,
  pkg: PackageAccessors,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const {partText, partBytes} = pkg;
  const drawingPath = sheetRelTarget(sheetPath, partText, 'drawing');
  if (drawingPath === undefined) return;
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
  pkg: PackageAccessors,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const mediaPath = sheetRelTarget(sheetPath, pkg.partText, 'image');
  if (mediaPath === undefined) return;
  let id = imageIdByMediaPath.get(mediaPath);
  if (id === undefined) {
    const bytes = pkg.partBytes(mediaPath);
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
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  sheet: Worksheet,
): void {
  const {partText, partBytes} = pkg;
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
    if (isPreservedSheetRelType(record.type)) capture(undefined, record.type, record.target);
  }
}

// A sheet relationship the model does not consume but must round-trip: a pivot table or a slicer. Every
// other sheet rel kind (drawing, printerSettings, table, comments, threadedComment, hyperlinks, background
// image, the comment VML) is modeled and re-serialised from the model, so preserving it here would emit
// the part twice.
function isPreservedSheetRelType(type: string): boolean {
  return type.endsWith('/pivotTable') || type.endsWith('/slicer');
}

// Capture the workbook-level references to package content the model does not interpret — pivot
// caches (`pivotCacheDefinition`), slicer caches (`slicerCache`), and external links (`externalLink`,
// each a link to a source workbook) — so a round-trip re-emits them instead of dropping the pivots,
// slicers, and linked-workbook references they back. A pivot cache's `<pivotCaches>` registration (its
// `cacheId`) and an external link's `<externalReferences>` position (its `[n]` index) are captured
// alongside so the wiring a pivot table or a formula resolves through survives too.
function readWorkbookPreservedReferences(
  workbookXml: string,
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const {partText, partBytes} = pkg;
  const relsXml = partText('xl/_rels/workbook.xml.rels');
  if (relsXml === undefined) return;
  const cacheIdByRelId = parsePivotCacheRegistrations(workbookXml);
  const externalIndexByRelId = parseExternalReferenceRegistrations(workbookXml);
  for (const record of parseRelationshipRecords(relsXml)) {
    if (record.external || !isPreservedWorkbookRelType(record.type)) continue;
    const entryPath = resolveWorkbookPart(record.target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts === undefined) continue;
    const cacheId = cacheIdByRelId.get(record.id);
    const externalReferenceIndex = externalIndexByRelId.get(record.id);
    workbook.addPreservedReference({
      relType: record.type,
      entryPath,
      parts,
      ...(cacheId !== undefined ? {pivotCacheId: cacheId} : {}),
      ...(externalReferenceIndex !== undefined ? {externalReferenceIndex} : {}),
    });
  }
}

// Content wired from the package's own `_rels/.rels` that the writer does not regenerate from the
// model — the ribbon customUI parts, custom document properties, a thumbnail. The writer rebuilds the
// root rels for the parts it models (the workbook, and core/app properties), so every other root
// relationship's target would be dropped on write; capturing its closure here re-declares it verbatim.
// External targets and the three regenerated relationship types are skipped.
function readRootPreservedReferences(
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const {partText, partBytes} = pkg;
  const relsXml = partText('_rels/.rels');
  if (relsXml === undefined) return;
  for (const record of parseRelationshipRecords(relsXml)) {
    if (record.external || isRegeneratedRootRelType(record.type)) continue;
    const entryPath = resolveRelativePart('', record.target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts === undefined) continue;
    workbook.addPreservedRootReference({relType: record.type, entryPath, parts});
  }
}

// The three root relationships the writer regenerates from the model on every write: the office
// document and the core/extended document properties. Every other root relationship is unmodeled and
// is preserved verbatim by {@link readRootPreservedReferences} rather than dropped.
function isRegeneratedRootRelType(type: string): boolean {
  return (
    type.endsWith('/officeDocument') ||
    type.endsWith('/core-properties') ||
    type.endsWith('/extended-properties')
  );
}

// A workbook relationship the model does not consume but must round-trip: a pivot cache, a slicer
// cache, an external link (the pointer to a linked source workbook), or a macro-enabled workbook's VBA
// project. Worksheets, styles, theme, shared strings, and the threaded-comment person registry are modeled
// and re-serialised from the model. Preserving vbaProject here — rather than silently dropping it, as an
// unrecognised relationship type otherwise would — is what keeps loading and re-saving a .xlsm from
// discarding its macros; the content-type override in workbook-xml.ts is the other half, so the re-emitted
// package still declares itself macro-enabled. Preserving externalLink is what keeps a formula's `[n]`
// external reference from dangling: the link part and its `<externalReferences>` registration are both
// re-emitted.
function isPreservedWorkbookRelType(type: string): boolean {
  return (
    type.endsWith('/pivotCacheDefinition') ||
    type.endsWith('/slicerCache') ||
    type.endsWith('/vbaProject') ||
    type.endsWith('/externalLink')
  );
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

// Map each `<externalReference>` in the workbook's `<externalReferences>` to its 0-based position, keyed
// by the relationship id it wires. That position is the `[n]` index a formula or defined name resolves
// an external cell through (`[1]Sheet!$A$1`), so preserving it lets the writer re-emit the block in the
// original order and keep every `[n]` pointing at the same linked workbook.
function parseExternalReferenceRegistrations(workbookXml: string): Map<string, number> {
  const byRelId = new Map<string, number>();
  let index = 0;
  for (const {attrs} of openElements(workbookXml, 'externalReference')) {
    if (attrs['r:id'] !== undefined) byRelId.set(attrs['r:id'], index++);
  }
  return byRelId;
}

// The `r:id` of the first `<drawing>` / `<legacyDrawingHF>` element in a worksheet, or undefined when
// the sheet declares none. The reference lives in the worksheet XML (not distinguishable by
// relationship Type — a header/footer VML and a comment VML share the `vmlDrawing` type), so the
// specific relationship is found by reading the element's `r:id` here.
function worksheetReferenceRelId(
  sheetXml: string,
  element: 'drawing' | 'legacyDrawingHF',
): string | undefined {
  for (const {attrs} of openElements(sheetXml, element)) {
    if (attrs['r:id'] !== undefined) return attrs['r:id'];
  }
  return undefined;
}

// A sheet's tables live in `xl/tables/table{n}.xml` parts, each reached through a relationship of
// type `.../table` on the sheet's own rels. The writer emits one relationship per table; each part
// is parsed back into the model and re-registered in definition order. A part that fails to parse
// (missing name/ref/columns — Excel corruption) is skipped rather than crashing the whole read.
function readSheetTables(sheetPath: string, pkg: PackageAccessors, sheet: Worksheet): void {
  const relsXml = pkg.partText(relsPathFor(sheetPath));
  if (relsXml === undefined) return;
  for (const target of relationshipTargetsByType(relsXml, 'table')) {
    const tableXml = pkg.partText(resolveRelativePart(sheetPath, target));
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
function readSheetPivotTables(sheetPath: string, pkg: PackageAccessors, sheet: Worksheet): void {
  const {partText} = pkg;
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
        credentials?: Partial<Record<WorkbookProtectionCredentialAttr, string>>;
      } = {};
      if (boolStrict(attrs.lockStructure)) protection.lockStructure = true;
      if (boolStrict(attrs.lockWindows)) protection.lockWindows = true;
      if (boolStrict(attrs.lockRevision)) protection.lockRevision = true;
      const credentials: Partial<Record<WorkbookProtectionCredentialAttr, string>> = {};
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

// Restore the workbook's saved window state from `<bookViews><workbookView/>` onto the model's view,
// so a round-trip hands back the geometry and active tab the author left rather than stamping the
// library's defaults over them. Only the first `<workbookView>` is read — the model carries one view,
// which is all Excel writes and all a single consuming window can restore.
//
// Each attribute is applied only when the source carried a usable value; an absent or non-numeric one
// leaves the default in place, so a truncated or hostile element degrades to a valid window rather
// than a NaN geometry that would serialise as garbage.
export function applyWorkbookView(view: WorkbookView, xml: string): void {
  for (const {attrs} of openElements(xml, 'workbookView')) {
    applyViewNumber(attrs.xWindow, (value) => (view.x = value));
    applyViewNumber(attrs.yWindow, (value) => (view.y = value));
    applyViewNumber(attrs.windowWidth, (value) => (view.width = value));
    applyViewNumber(attrs.windowHeight, (value) => (view.height = value));
    applyViewNumber(attrs.activeTab, (value) => (view.activeTab = value));
    if (attrs.visibility === 'hidden' || attrs.visibility === 'veryHidden') {
      view.visibility = attrs.visibility;
    }
    if (boolStrict(attrs.minimized)) view.minimized = true;
    return;
  }
}

function applyViewNumber(raw: string | undefined, assign: (value: number) => void): void {
  if (raw === undefined) return;
  const value = Number(raw);
  if (Number.isFinite(value)) assign(Math.trunc(value));
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
