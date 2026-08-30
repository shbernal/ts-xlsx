// The buffered `.xlsx` reader: an OPC zip package in, a Workbook model out.
//
// It reconstructs the part of the model the writer emits today: sheet names and order,
// cells holding a number, string, boolean, or formula, per-column width/visibility,
// per-row height/visibility, merged ranges, page margins, and cell styles (pattern fills,
// number formats, fonts, borders, alignment, and protection, per cell or inherited from a
// formatted row/column). Shared-formula slaves and the richer value kinds land as the model
// grows; an unrecognised construct is skipped rather than guessed, so a foreign file reads
// without crashing even where a facet is not yet materialised.
//
// This module is the orchestrator. It wires the parsed package parts together (the OPC/rel
// resolution (`./read-opc.ts`), the style table (`./read-styles.ts`), and each worksheet body
// (`./read-worksheet.ts`)) and owns the sheet-part discovery (notes, images, tables, pivots) and
// preserved-reference capture that a faithful round-trip depends on.
//
// Untrusted input: inflate is bounded by a running byte counter (`./inflate.ts`) that caps
// actual decompressed output rather than trusting the archive's forgeable size headers, and
// the parser (ADR 0004) never expands entities.

import {decodeRange} from '../../core/address.ts';
import type {CommentThread} from '../../core/comment-thread.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {INTERNAL} from '../../core/internal.ts';
import type {PreservedWorksheetReference} from '../../core/preserved.ts';
import {
  WORKBOOK_PROTECTION_CREDENTIAL_ATTRS,
  type WorkbookProtection,
  type WorkbookProtectionCredentialAttr,
} from '../../core/workbook-protection.ts';
import {type DefinedName, Workbook, type WorkbookView} from '../../core/workbook.ts';
import {isVisibility, type Worksheet, type WorksheetState} from '../../core/worksheet.ts';
import {
  boolStrict,
  capturedText,
  enumToken,
  localName,
  numInteger,
  openElements,
  parseXml,
  parseXmlPasses,
} from '../../xml/xml-read.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {extensionOf} from '../opc/part-paths.ts';
import {
  capturePartClosure,
  contentTypeResolver,
  type PackageAccessors,
  type PartRelationships,
  openSpreadsheetPackage,
  parseRelationshipRecords,
  readPartRelationships,
  resolveRelativePart,
} from '../opc/read-opc.ts';
import type {ReadXlsxOptions} from '../opc/read-options.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {readXlsbPackage, XLSB_WORKBOOK_PART} from '../xlsb/read.ts';
import type {SharedString} from './cell-value.ts';
import {applyNotes, type ParsedComment, parseComments} from './comments.ts';
import {conditionalFormattingPass} from './conditional-formatting.ts';
import {
  applyDataValidations,
  dataValidationPass,
  extendedDataValidationPass,
} from './data-validation.ts';
import {applyHyperlinks, sheetHyperlinkPass} from './hyperlinks.ts';
import {drawingHasUnmodeledContent, parseDrawing} from './images.ts';
import {parsePivotTable} from './read-pivot.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseStyleTable} from './read-styles.ts';
import {worksheetPass} from './read-worksheet.ts';
import {parseTable} from './tables.ts';
import {parseThemeColorScheme, parseThemeFontScheme} from './theme-xml.ts';
import {buildCommentThreads, parsePersons, parseThreadedComments} from './threaded-comments.ts';

// The read option bag is shared with the `.xlsb` reader and the row streamer, so it is declared apart
// from all three; it stays reachable here because this is the entry point callers reach for. The
// bound's default is not re-exported: `openSpreadsheetPackage` applies it, and no caller names it.
export type {ReadXlsxOptions} from '../opc/read-options.ts';
export type {StyleTable, XfStyle} from '../style/xf-style.ts';
export {parseStyleTable} from './read-styles.ts';

/**
 * Read a spreadsheet package into a {@link Workbook}.
 *
 * Both OOXML serialisations are accepted: an XML `.xlsx`, and a binary `.xlsb` (BIFF12), which is the
 * same OPC container with binary office-document parts. The two are auto-detected from the package
 * itself rather than from a file extension, so a caller never branches on which form it holds, and
 * the model produced is the same either way. See `../xlsb/read.ts` for what the binary path does not
 * yet decode.
 *
 * @throws {UnsupportedFormatError} if the input is neither: a legacy `.xls` (`.format === 'xls'`) or
 *   an unrecognised/non-ZIP blob (`'unknown'`).
 * @throws {XlsbParseError} if a binary `.xlsb` part is malformed.
 * @throws {PackageReadError} if the input is a ZIP that cannot be unpacked: a corrupt or
 *   truncated archive, or one exceeding the inflate bound (a probable zip bomb).
 */
export function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook {
  const {files, pkg, workbookXml} = openSpreadsheetPackage(data, options.maxUncompressedBytes);
  const {partText} = pkg;

  if (workbookXml === undefined) {
    // No XML office document. A binary one means this is an `.xlsb`, which reads through the BIFF12
    // codec over the very same model. The package is already inflated, so it is handed over as-is.
    if (files[XLSB_WORKBOOK_PART] !== undefined) return readXlsbPackage(files);
    throw new UnsupportedFormatError('unknown');
  }

  // A part's content type is needed to faithfully re-declare any part preserved verbatim for
  // round-tripping (a vector-shape drawing, a header/footer image and its VML). Resolve it the way
  // OPC does: an explicit `<Override>` for the exact part, else the `<Default>` for its extension.
  const contentTypeOf = contentTypeResolver(partText('[Content_Types].xml') ?? '');

  // One parse of the workbook's rels, queried by the sheet loop and by the two workbook-level part
  // readers below. It used to be held as a raw string and handed to three separate scanners, which
  // also left two different idioms for "reach a related part" side by side in one function.
  const workbookRels = readPartRelationships('xl/workbook.xml', partText);
  const sharedStrings = parseSharedStrings(partText('xl/sharedStrings.xml') ?? '');
  // The style table resolves a cell/row/column style index to its facets (fill, number
  // format); a package without one (a hand-rolled foreign file) yields an empty table and
  // every index reads as unstyled.
  const stylesXml = partText('xl/styles.xml') ?? '';
  const {cellXfs: xfStyles, namedStyles, defaultFont, preserved} = parseStyleTable(stylesXml);

  const workbook = new Workbook();
  // The four sub-tables the stylesheet carries verbatim, all captured by the same read of the part
  // that resolved the xfs above rather than by four more scans of it.
  //
  // Preserve the differential-style table so conditional formatting's dxfId references stay valid,
  // and a foreign dxf's number format stays a real format code, across a re-write.
  workbook[INTERNAL].restoreDifferentialStyles([...preserved.dxfs]);
  // Preserve a custom indexed-color palette so an `indexed="…"` colour reference keeps its intended
  // RGB across a re-write instead of resolving to a different default-palette entry.
  workbook[INTERNAL].restoreIndexedColors([...preserved.indexedColors]);
  // Preserve the author's "Recent Colors" swatches, which the model never reads but re-writing would
  // otherwise discard.
  workbook[INTERNAL].restoreMruColors([...preserved.mruColors]);
  // Preserve the custom table-style definitions so a table referencing one by name still resolves to
  // a real definition after a re-write instead of rendering unstyled.
  workbook[INTERNAL].restoreTableStyles(preserved.tableStyles);
  // Preserve the theme part so a branded colour/font scheme is not overwritten by the default theme
  // the writer emits for a workbook that has none.
  readWorkbookTheme(workbookRels, pkg, contentTypeOf, workbook);
  // Preserve the named cell-style layer only when a file declares one beyond the Normal default, so an
  // ordinary workbook keeps an empty named-style table and emits just the default on write.
  if (namedStyles.length > 1) workbook[INTERNAL].restoreNamedStyles(namedStyles);
  // Preserve the declared default font (font id 0) so a re-write emits the face the file itself named
  // rather than an assumed Calibri, which would change every empty cell and the metric every
  // character-unit column width is expressed in.
  workbook[INTERNAL].restoreDefaultFont(defaultFont);
  const core = partText('docProps/core.xml');
  if (core !== undefined) applyCoreProperties(workbook, core);
  const app = partText('docProps/app.xml');
  if (app !== undefined) applyAppProperties(workbook, app);
  workbook.protection = parseWorkbookProtection(workbookXml);
  applyWorkbookView(workbook.view, workbookXml);
  // The threaded-comment author registry is workbook-level, and every conversation on every sheet
  // resolves its authors and @mentions through it, so it is restored before the sheet loop that reads
  // those conversations, not alongside the other workbook-level parts below.
  readWorkbookPersons(workbookRels, pkg, workbook);

  const context: SheetReadContext = {
    pkg,
    workbook,
    contentTypeOf,
    sharedStrings,
    xfStyles,
    // A picture used on more than one sheet is one media part; caching by media path across the
    // whole loop keeps it a single workbook image so a re-write does not duplicate the bytes.
    imageIdByMediaPath: new Map<string, number>(),
  };
  const sheetOrder: string[] = [];
  for (const {name, relId, state} of parseWorkbookSheets(workbookXml)) {
    const target = workbookRels.byId(relId)?.target;
    const sheet = workbook.addWorksheet(name, state === undefined ? undefined : {state});
    sheetOrder.push(name);
    readSheet(sheet, target === undefined ? undefined : workbookRels.pathOf(target), context);
  }

  readWorkbookPreservedReferences(workbookXml, workbookRels, pkg, contentTypeOf, workbook);
  readRootPreservedReferences(pkg, contentTypeOf, workbook);

  // Defined names follow the sheets: a scoped name's `localSheetId` indexes the sheet order, which
  // is why the names are read only once every sheet is registered.
  for (const name of parseWorkbookDefinedNames(workbookXml, sheetOrder)) {
    workbook.defineName(name);
  }
  return workbook;
}

/**
 * Everything a single sheet needs from the package around it, gathered once for the whole sheet loop
 * so {@link readSheet} takes a context rather than seven positional arguments. `imageIdByMediaPath`
 * is the one mutable member, and is deliberately shared across sheets: that sharing is what makes a
 * picture used on two of them resolve to one workbook image rather than two copies of the bytes.
 */
interface SheetReadContext {
  readonly pkg: PackageAccessors;
  readonly workbook: Workbook;
  readonly contentTypeOf: (path: string) => string;
  readonly sharedStrings: readonly SharedString[];
  readonly xfStyles: readonly XfStyle[];
  readonly imageIdByMediaPath: Map<string, number>;
}

/**
 * Read one worksheet at `path`: its body, the four overlays that ride the same parse of the worksheet
 * part, and every part hanging off the sheet's own relationships.
 *
 * The stages are ordered, not merely sequential, and each constraint is non-local:
 *
 * - the overlays are gathered during the body's parse but *applied* only once the sheet's rels are in
 *   hand, because a hyperlink resolves its target through them;
 * - threads land before notes, because a threaded cell's comments-part entry is that thread's legacy
 *   fallback rather than a note, and `applyNotes` reads the restored threads to tell the two apart;
 * - preserved references are captured after the images, because that capture excludes what the image
 *   reader already modelled and would otherwise re-emit a drawing the writer also emits.
 *
 * Defined names are deliberately *not* read here: a sheet-scoped name indexes the workbook's sheet
 * order, so `readXlsx` reads them only once every sheet is registered.
 *
 * A sheet whose relationship is dangling (`path === undefined`) stays an empty sheet in its place in
 * the order rather than vanishing from the workbook.
 */
function readSheet(sheet: Worksheet, path: string | undefined, context: SheetReadContext): void {
  const {pkg, workbook, contentTypeOf, sharedStrings, xfStyles, imageIdByMediaPath} = context;
  const {partText} = pkg;
  const sheetXml = path === undefined ? undefined : partText(path);

  // Five readers want the worksheet part, and it is the largest in the package by a wide margin, so
  // they share one parse of it rather than scanning it once each. Only the body commits as it goes;
  // the other four gather, and are applied below in the order they were always applied.
  const hyperlinks = sheetHyperlinkPass();
  const validations = dataValidationPass();
  const extendedValidations = extendedDataValidationPass();
  const formattings = conditionalFormattingPass();
  if (sheetXml !== undefined) {
    parseXmlPasses(sheetXml, [
      worksheetPass(sheet, sharedStrings, xfStyles),
      hyperlinks,
      validations,
      extendedValidations,
      formattings,
    ]);
  }
  if (path === undefined) return;

  // The sheet's rels are the index to nearly every part hanging off it, so they are parsed once here
  // and threaded through the readers below rather than re-read by each.
  const sheetRels = readPartRelationships(path, partText);
  if (sheetXml !== undefined) {
    applyHyperlinks(sheet, hyperlinks.result(), (id) => sheetRels.byId(id)?.target);
    applyDataValidations(sheet, [...validations.result(), ...extendedValidations.result()]);
    for (const cf of formattings.result()) sheet.addConditionalFormatting(cf);
  }

  const threads = readSheetCommentThreads(sheetRels, pkg, workbook);
  if (threads.length > 0) sheet[INTERNAL].restoreCommentThreads(threads);
  const comments = readSheetComments(sheetRels, pkg);
  if (comments !== undefined) applyNotes(sheet, comments);

  readSheetImages(sheetRels, pkg, workbook, sheet, imageIdByMediaPath);
  readSheetBackground(sheetRels, pkg, workbook, sheet, imageIdByMediaPath);
  if (sheetXml !== undefined) {
    readSheetPreservedReferences(sheetRels, sheetXml, pkg, contentTypeOf, sheet);
  }

  readSheetTables(sheetRels, pkg, sheet);
  readSheetPivotTables(sheetRels, pkg, sheet);
  const printerSettings = readSheetPrinterSettings(sheetRels, pkg);
  if (printerSettings !== undefined) sheet.pageSetup.printerSettings = printerSettings;
}

// A sheet's comments live in a comments part reached through the sheet's own relationships: the sheet
// declares a relationship of type `.../comments` whose target resolves (relative to the sheet's
// directory) to the comments part. A sheet declaring no such relationship simply has none.
function readSheetComments(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
): Map<string, ParsedComment> | undefined {
  const commentsPath = sheetRels.targetPath('comments');
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
  workbookRels: PartRelationships,
  pkg: PackageAccessors,
  workbook: Workbook,
): void {
  const path = workbookRels.targetPath('person');
  const xml = path === undefined ? undefined : pkg.partText(path);
  if (xml !== undefined) workbook[INTERNAL].restorePersons(parsePersons(xml));
}

// The workbook's theme part: the `<clrScheme>`/`<fontScheme>`/`<fmtScheme>` every `theme="n"` colour
// and every `scheme="major|minor"` font in the package resolves against. It is reached through the
// workbook's `.../theme` relationship rather than assumed at `xl/theme/theme1.xml`, because the target
// is rel-relative and a foreign package is free to name the part anything.
//
// Captured with its transitive part closure, not as a lone string: a theme can carry its own
// relationships (a picture used as a themed fill, wired by an `r:embed` into the theme's rels part),
// and re-emitting the theme body without them would leave that reference dangling, which Excel
// reports as a package needing repair. A package that declares no theme leaves the workbook on the
// library's default, which is also what a dangling relationship target degrades to.
function readWorkbookTheme(
  workbookRels: PartRelationships,
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const entryPath = workbookRels.targetPath('theme');
  if (entryPath === undefined) return;
  const parts = capturePartClosure(entryPath, pkg.partText, pkg.partBytes, contentTypeOf);
  if (parts === undefined) return;
  // The schemes are decoded here rather than on demand from the model: the part rides through the
  // model as opaque bytes, and only the codec knows how to read one.
  const xml = pkg.partText(entryPath) ?? '';
  workbook[INTERNAL].restoreThemePart(
    {entryPath, parts},
    {colors: parseThemeColorScheme(xml), fonts: parseThemeFontScheme(xml)},
  );
}

// A sheet's threaded conversations live in a `xl/threadedComments/threadedComment{n}.xml` part reached
// through a relationship of type `.../threadedComment` on the sheet's own rels: the same discovery
// shape as the notes part above, and deliberately separate from it: a thread and a legacy note are
// different features that happen to share a sheet. The messages are grouped into threads and their
// authors resolved against the workbook registry, so each thread lands self-contained.
//
// What lands here IS what a re-write emits: the thread part is re-serialised from these threads, and so is
// the legacy fallback `<comment>` that binds each cell to its conversation (see `comments.ts`). Anything
// this reader drops is therefore dropped from the file, which is why a message too damaged to place is
// still kept wherever it can be, and why the anchor is canonicalised here rather than trusted downstream.
function readSheetCommentThreads(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
  workbook: Workbook,
): CommentThread[] {
  const path = sheetRels.targetPath('threadedComment');
  const xml = path === undefined ? undefined : pkg.partText(path);
  if (xml === undefined) return [];
  return buildCommentThreads(parseThreadedComments(xml), (id) => workbook.getPerson(id));
}

// A sheet's printer-settings blob is an opaque binary part linked from `<pageSetup r:id>`: the sheet
// declares a relationship of type `.../printerSettings` whose target resolves to a `.bin` part. We
// keep the raw bytes verbatim: the DEVMODE inside is platform-specific and the model never
// interprets it, only round-trips it so re-writing the file preserves the user's print configuration.
// A sheet declaring no such relationship simply has none.
function readSheetPrinterSettings(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
): Uint8Array | undefined {
  const path = sheetRels.targetPath('printerSettings');
  return path === undefined ? undefined : pkg.partBytes(path);
}

// One workbook image per media part, however many places in the package point at that part. A sheet's
// background and a drawing's picture routinely name the same bytes, and modelling them as two images
// would write the media twice on the way out, so the map is the property both callers depend on and
// is what makes the re-write byte-count-stable. `undefined` means the part named a target the package
// does not carry, which is a broken relationship the caller skips over rather than fails on.
function internImage(
  workbook: Workbook,
  pkg: PackageAccessors,
  imageIdByMediaPath: Map<string, number>,
  mediaPath: string,
): number | undefined {
  const known = imageIdByMediaPath.get(mediaPath);
  if (known !== undefined) return known;
  const bytes = pkg.partBytes(mediaPath);
  if (bytes === undefined) return undefined;
  const id = workbook.addImage({buffer: bytes, extension: extensionOf(mediaPath)});
  imageIdByMediaPath.set(mediaPath, id);
  return id;
}

// A sheet's anchored images live in a drawing part reached through the sheet's own relationships: a
// relationship of type `.../drawing` names the drawing part, whose own relationships map each
// picture's embed id to a media part under `xl/media/`. Each anchor becomes a workbook image (deduped
// by media path) placed back on the sheet at its two-cell anchor.
function readSheetImages(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const {partText} = pkg;
  const drawingPath = sheetRels.targetPath('drawing');
  if (drawingPath === undefined) return;
  const drawingXml = partText(drawingPath);
  if (drawingXml === undefined) return;
  // A drawing that also holds a chart or shape is preserved whole (see readSheetPreservedReferences),
  // so its pictures must not be modeled here: modeling them would leave the sheet with images, which
  // suppresses that preservation and drops the chart. Leaving `sheet.images` empty routes the entire
  // drawing, pictures included, through byte-preservation, keeping every anchor faithful.
  if (drawingHasUnmodeledContent(drawingXml)) return;
  const drawingRels = readPartRelationships(drawingPath, partText);

  for (const anchor of parseDrawing(drawingXml)) {
    const embedded = drawingRels.byId(anchor.embed);
    if (embedded === undefined) continue;
    const mediaPath = drawingRels.pathOf(embedded.target);
    const id = internImage(workbook, pkg, imageIdByMediaPath, mediaPath);
    if (id === undefined) continue;
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
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
  workbook: Workbook,
  sheet: Worksheet,
  imageIdByMediaPath: Map<string, number>,
): void {
  const mediaPath = sheetRels.targetPath('image');
  if (mediaPath === undefined) return;
  const id = internImage(workbook, pkg, imageIdByMediaPath, mediaPath);
  if (id !== undefined) sheet.addBackgroundImage(id);
}

// Capture the worksheet-level references to package content the model does not interpret, so a
// round-trip re-emits them verbatim instead of dropping them:
//   • `<drawing>`, but only when the reader modeled no anchored image from it: either a drawing that
//     holds no pictures at all (a chart or shape), or a mixed drawing whose pictures the reader
//     declined to model precisely so the whole part (chart included) rides here verbatim. A drawing
//     whose pictures were modeled is owned by the model and re-serialised from it; capturing it here
//     too would double-emit those pictures.
//   • `<legacyDrawingHF>`, a header/footer image's VML, which the model never interprets.
// Each reference's target part and the transitive closure of parts it reaches (a VML's image, a
// drawing's media) are captured with their bytes, content types, and relationships.
function readSheetPreservedReferences(
  sheetRels: PartRelationships,
  sheetXml: string,
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  sheet: Worksheet,
): void {
  const {partText, partBytes} = pkg;

  const capture = (
    element: PreservedWorksheetReference['element'],
    relType: string,
    target: string,
  ): void => {
    const entryPath = sheetRels.pathOf(target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts !== undefined)
      sheet[INTERNAL].addPreservedReference({element, relType, entryPath, parts});
  };

  // Element-wired references: a `<drawing>`/`<legacyDrawingHF>` names its part by an `r:id` in the
  // sheet body. A `<drawing>` is preserved only when the reader modeled no picture from it (a
  // chart/shape-only drawing, or a mixed one the reader left unmodeled) since one whose pictures are
  // modeled is re-serialised from the model.
  const referenceElements: Array<'drawing' | 'legacyDrawingHF'> =
    sheet.images.length === 0 ? ['drawing', 'legacyDrawingHF'] : ['legacyDrawingHF'];
  for (const element of referenceElements) {
    const relId = worksheetReferenceRelId(sheetXml, element);
    const record = relId === undefined ? undefined : sheetRels.byId(relId);
    if (record !== undefined && !record.external) capture(element, record.type, record.target);
  }

  // Relationship-wired references: a pivot table or slicer is reached through a sheet relationship
  // with no worksheet child pointing at it; Excel discovers it by scanning the sheet's rels. Preserve
  // each so the pivots/slicers a fill-and-save workflow does not touch are not dropped.
  for (const record of sheetRels.records) {
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

// Capture the workbook-level references to package content the model does not interpret: pivot
// caches (`pivotCacheDefinition`), slicer caches (`slicerCache`), and external links (`externalLink`,
// each a link to a source workbook), so a round-trip re-emits them instead of dropping the pivots,
// slicers, and linked-workbook references they back. A pivot cache's `<pivotCaches>` registration (its
// `cacheId`) and an external link's `<externalReferences>` position (its `[n]` index) are captured
// alongside so the wiring a pivot table or a formula resolves through survives too.
function readWorkbookPreservedReferences(
  workbookXml: string,
  workbookRels: PartRelationships,
  pkg: PackageAccessors,
  contentTypeOf: (path: string) => string,
  workbook: Workbook,
): void {
  const {partText, partBytes} = pkg;
  const cacheIdByRelId = parsePivotCacheRegistrations(workbookXml);
  const externalIndexByRelId = parseExternalReferenceRegistrations(workbookXml);
  for (const record of workbookRels.records) {
    if (record.external || !isPreservedWorkbookRelType(record.type)) continue;
    const entryPath = workbookRels.pathOf(record.target);
    const parts = capturePartClosure(entryPath, partText, partBytes, contentTypeOf);
    if (parts === undefined) continue;
    const cacheId = cacheIdByRelId.get(record.id);
    const externalReferenceIndex = externalIndexByRelId.get(record.id);
    workbook[INTERNAL].addPreservedReference({
      relType: record.type,
      entryPath,
      parts,
      ...(cacheId !== undefined ? {pivotCacheId: cacheId} : {}),
      ...(externalReferenceIndex !== undefined ? {externalReferenceIndex} : {}),
    });
  }
}

// Content wired from the package's own `_rels/.rels` that the writer does not regenerate from the
// model: the ribbon customUI parts, custom document properties, a thumbnail. The writer rebuilds the
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
    workbook[INTERNAL].addPreservedRootReference({relType: record.type, entryPath, parts});
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
// and re-serialised from the model. Preserving vbaProject here, rather than silently dropping it as an
// unrecognised relationship type otherwise would, is what keeps loading and re-saving a .xlsm from
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
// relationship Type, since a header/footer VML and a comment VML share the `vmlDrawing` type), so the
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
// (missing name/ref/columns, which is Excel corruption) is skipped rather than crashing the whole read.
function readSheetTables(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
  sheet: Worksheet,
): void {
  for (const tablePath of sheetRels.targetPaths('table')) {
    const tableXml = pkg.partText(tablePath);
    if (tableXml === undefined) continue;
    const options = parseTable(tableXml);
    if (options !== undefined) sheet.addTable(options);
  }
  dropMergesInsideTables(sheet);
}

// Reconstruct an inspectable model of each pivot table hosted on a sheet. A pivot is reached by a
// sheet relationship of type `.../pivotTable`; the pivot-table part carries its own relationship of
// type `.../pivotCacheDefinition` to the cache holding the field catalogue and source range. Both
// parts are parsed and combined into a read-only view registered on the sheet, separate from the
// byte-preservation that actually round-trips the pivot, so this never changes what is re-emitted.
// The read is lenient: a pivot whose cache is missing still yields a (partial) model rather than
// throwing, matching Excel's tolerance for a damaged package on load.
function readSheetPivotTables(
  sheetRels: PartRelationships,
  pkg: PackageAccessors,
  sheet: Worksheet,
): void {
  const {partText} = pkg;
  for (const tablePath of sheetRels.targetPaths('pivotTable')) {
    const tableXml = partText(tablePath);
    if (tableXml === undefined) continue;
    const cachePath = readPartRelationships(tablePath, partText).targetPath('pivotCacheDefinition');
    const cacheXml = cachePath === undefined ? '' : (partText(cachePath) ?? '');
    sheet[INTERNAL].addLoadedPivotTable(parsePivotTable(tableXml, cacheXml));
  }
}

// Excel forbids a merged range inside a formatted table and repairs such a file on load by dropping
// the merge. A worksheet's merges are read before its tables, so a real file carrying that invalid
// geometry lands in the model intact; this applies the same repair once the tables are known, so a
// re-write does not surface the Excel-invalid geometry the writer (correctly) rejects.
function dropMergesInsideTables(sheet: Worksheet): void {
  const regions = sheet.tables.map((table) => table.region);
  if (regions.length === 0) return;
  // The copy is not incidental: `sheet.merges` is the live backing array and `unmergeCells` splices
  // out of it, so iterating it directly would skip the entry after every removal.
  // oxlint-disable-next-line unicorn/no-useless-spread
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
    // `visible` is the schema default and the model's, so it is dropped rather than stored: keeping
    // it would put a `state="visible"` attribute into a file Excel writes without one.
    const state = enumToken(attrs.state, isVisibility);
    if (state !== undefined && state !== 'visible') entry.state = state;
    sheets.push(entry);
  }
  return sheets;
}

// Read the workbook's structure/window protection (`<workbookProtection>`). The three lock flags are
// decoded as booleans (an absent or "0" attribute stays unlocked), and only the whitelisted
// password/agile-hash attributes are preserved verbatim: a hostile or unknown attribute is dropped
// rather than echoed back on write. Returns undefined when the workbook declares no protection.
function parseWorkbookProtection(xml: string): WorkbookProtection | undefined {
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
// library's defaults over them. Only the first `<workbookView>` is read: the model carries one view,
// which is all Excel writes and all a single consuming window can restore.
//
// Each attribute is applied only when the source carried a usable value; an absent or non-numeric one
// leaves the default in place, so a truncated or hostile element degrades to a valid window rather
// than a NaN geometry that would serialise as garbage.
export function applyWorkbookView(view: WorkbookView, xml: string): void {
  for (const {attrs} of openElements(xml, 'workbookView')) {
    // The window may sit at a negative origin (a secondary monitor left of the primary), so only
    // the extents and the tab ordinal carry a floor.
    const x = numInteger(attrs.xWindow);
    if (x !== undefined) view.x = x;
    const y = numInteger(attrs.yWindow);
    if (y !== undefined) view.y = y;
    const width = numInteger(attrs.windowWidth, 0);
    if (width !== undefined) view.width = width;
    const height = numInteger(attrs.windowHeight, 0);
    if (height !== undefined) view.height = height;
    const activeTab = numInteger(attrs.activeTab, 0);
    if (activeTab !== undefined) view.activeTab = activeTab;
    const visibility = enumToken(attrs.visibility, isVisibility);
    if (visibility !== undefined && visibility !== 'visible') view.visibility = visibility;
    if (boolStrict(attrs.minimized)) view.minimized = true;
    return;
  }
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
      const scopeIndex = numInteger(attrs.localSheetId, 0) ?? -1;
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
const CORE_PROPERTY_LOCAL_NAMES = new Set([
  'title',
  'creator',
  'lastModifiedBy',
  'created',
  'modified',
]);

function applyCoreProperties(workbook: Workbook, xml: string): void {
  for (const {local, text} of capturedText(xml, CORE_PROPERTY_LOCAL_NAMES)) {
    if (local === 'title') workbook.properties.title = text;
    else if (local === 'creator') workbook.properties.creator = text;
    else if (local === 'lastModifiedBy') workbook.properties.lastModifiedBy = text;
    else {
      // An unparseable date is dropped rather than stored as an Invalid Date, which would write
      // back as the string `Invalid Date` and lose the property for good.
      const date = new Date(text);
      if (!Number.isNaN(date.getTime())) {
        if (local === 'created') workbook.properties.created = date;
        else workbook.properties.modified = date;
      }
    }
  }
}

// `Company` is the one document property OOXML keeps in the extended part rather than the core
// one. Everything else in app.xml is either derived (`TitlesOfParts`) or this library's own
// (`Application`), so nothing here reads more than the single element.
function applyAppProperties(workbook: Workbook, xml: string): void {
  for (const {text} of capturedText(xml, 'Company')) workbook.properties.company = text;
}
