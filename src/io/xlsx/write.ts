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
import type {Cell} from '../../core/cell.ts';
import {dateToSerial, DEFAULT_DATE_NUMFMT} from '../../core/date.ts';
import type {WorkbookImage} from '../../core/image.ts';
import {mangleFormula} from '../../core/formula.ts';
import type {Table, TableColumn} from '../../core/table.ts';
import {
  detectValueType,
  type FormulaResult,
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
  PageMargins,
  PageSetup,
  RowProperties,
  Worksheet,
  WorksheetProperties,
} from '../../core/worksheet.ts';
import {collectNotes, commentsXml, type NoteCell, vmlDrawingXml} from './comments.ts';
import {
  collectHyperlinks,
  hyperlinksXml,
  planHyperlinks,
  type PlannedHyperlink,
} from './hyperlinks.ts';
import {type DrawingImage, drawingRelsXml, drawingXml, imageContentType} from './images.ts';
import {richTextRunsXml} from './rich-text.ts';
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
  hyperlink: `${NS.docRels}/hyperlink`,
} as const;

/**
 * Serialise a workbook into an `.xlsx` package.
 *
 * @throws {Error} if the workbook has no worksheets (a zero-sheet package is corrupt),
 *   or holds a value the writer cannot yet represent.
 */
export function writeXlsx(workbook: Workbook): Uint8Array {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) {
    throw new Error('cannot write a workbook with no worksheets — a zero-sheet package is corrupt to Excel');
  }

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

  // Serialise the worksheets first: interning each cell/row fill into the style table is a
  // side effect of that pass, so styles.xml can only be generated once every sheet is done.
  const styles = new StyleRegistry();
  const sheetXml = sheets.map((sheet, i) =>
    worksheetXml(
      sheet,
      sheetTables[i] ?? [],
      styles,
      sheetDrawings[i]?.relId ?? null,
      sheetComments[i]?.vmlRelId ?? null,
      sheetPrinterSettings[i]?.relId ?? null,
      sheetHyperlinks[i] ?? []
    )
  );

  const commentNumbers = sheetComments.filter((c): c is CommentPlan => c !== null).map(c => c.number);
  const drawingNumbers = sheetDrawings.filter((d): d is DrawingPlan => d !== null).map(d => d.number);
  const printerSettingsNumbers = sheetPrinterSettings
    .filter((p): p is PrinterSettingsPlan => p !== null)
    .map(p => p.number);
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      contentTypesXml(
        sheets.length,
        allTables,
        commentNumbers,
        drawingNumbers,
        printerSettingsNumbers,
        media.extensions
      )
    ),
    '_rels/.rels': strToU8(rootRelsXml()),
    'docProps/core.xml': strToU8(corePropsXml(workbook.properties)),
    'docProps/app.xml': strToU8(appPropsXml()),
    'xl/workbook.xml': strToU8(workbookXml(workbook)),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelsXml(sheets.length)),
    'xl/styles.xml': strToU8(styles.toXml()),
    'xl/theme/theme1.xml': strToU8(THEME1_XML),
  };
  for (const part of media.parts) {
    files[`xl/media/image${part.number}.${part.extension}`] = part.data;
  }
  sheets.forEach((_sheet, i) => {
    const tables = sheetTables[i] ?? [];
    const drawing = sheetDrawings[i] ?? null;
    const comments = sheetComments[i] ?? null;
    const printerSettings = sheetPrinterSettings[i] ?? null;
    const hyperlinks = sheetHyperlinks[i] ?? [];
    const hasExternalHyperlink = hyperlinks.some((link) => link.relId !== undefined);
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml[i] as string);
    if (
      tables.length > 0 ||
      drawing !== null ||
      comments !== null ||
      printerSettings !== null ||
      hasExternalHyperlink
    ) {
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(
        worksheetRelsXml(tables, drawing, comments, printerSettings, hyperlinks)
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

  return zipSync(files, {level: 6});
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

// Gather the workbook images actually anchored by some sheet (an unreferenced image is not written),
// number them in first-use order, and record the extensions in play. A sheet anchoring an id with no
// registered image is a programming error the writer surfaces rather than emitting a dangling embed.
function planMedia(workbook: Workbook, sheets: readonly Worksheet[]): MediaPlan {
  const usedIds: number[] = [];
  const seen = new Set<number>();
  for (const sheet of sheets) {
    for (const image of sheet.images) {
      if (!seen.has(image.imageId)) {
        seen.add(image.imageId);
        usedIds.push(image.imageId);
      }
    }
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

function contentTypesXml(
  sheetCount: number,
  tables: readonly PlannedTable[],
  commentNumbers: readonly number[],
  drawingNumbers: readonly number[],
  printerSettingsNumbers: readonly number[],
  mediaExtensions: readonly string[]
): string {
  const overrides = [
    override('/xl/workbook.xml', CT.workbook),
    ...range(sheetCount).map(i => override(`/xl/worksheets/sheet${i + 1}.xml`, CT.worksheet)),
    ...tables.map(({number}) => override(`/xl/tables/table${number}.xml`, CT.table)),
    ...drawingNumbers.map(number => override(`/xl/drawings/drawing${number}.xml`, CT.drawing)),
    ...commentNumbers.map(number => override(`/xl/comments${number}.xml`, CT.comments)),
    override('/xl/theme/theme1.xml', CT.theme),
    override('/xl/styles.xml', CT.styles),
    override('/docProps/core.xml', CT.core),
    override('/docProps/app.xml', CT.app),
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

function workbookXml(workbook: Workbook): string {
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
    '</workbook>'
  );
}

// The `<definedNames>` block follows `<sheets>` in the schema. A sheet-scoped name carries a
// `localSheetId` — the 0-based position of its sheet among the `<sheet>` entries, NOT the sheet's
// own id — so the index is resolved against the worksheet order here. The refersTo formula is the
// element's text content, run through the same `_xlfn.` function mangling the writer applies to a
// cell formula so a name defined as a modern function (a LAMBDA, an XLOOKUP-based name) is stored
// under the prefix Excel requires; a plain reference has no function call and passes through
// untouched. Only names that are actually set emit anything.
function definedNamesXml(workbook: Workbook): string {
  const names = workbook.definedNames;
  if (names.length === 0) return '';
  const sheets = workbook.worksheets;
  const entries = names
    .map(name => {
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
    })
    .join('');
  return `<definedNames>${entries}</definedNames>`;
}

function workbookRelsXml(sheetCount: number): string {
  const rels = [
    ...range(sheetCount).map(i =>
      relationship(`rId${i + 1}`, REL.worksheet, `worksheets/sheet${i + 1}.xml`)
    ),
    relationship(`rId${sheetCount + 1}`, REL.styles, 'styles.xml'),
    relationship(`rId${sheetCount + 2}`, REL.theme, 'theme/theme1.xml'),
  ].join('');
  return XML_DECLARATION + `<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

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
  hyperlinks: readonly PlannedHyperlink[]
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
        });
        return cellXml(cell, style, sharedRoles.get(cell.address));
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
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    sheetFormatPr(sheet.properties) +
    colsXml(sheet, styles) +
    sheetData +
    sheetProtectionXml(sheet.protection) +
    mergeCellsXml(sheet.merges) +
    // CT_Worksheet order: <hyperlinks> follows <mergeCells> (and would follow any conditional
    // formatting / data validations) and precedes the print settings.
    hyperlinksXml(hyperlinks) +
    pageMarginsXml(sheet.pageMargins) +
    pageSetupXml(sheet.pageSetup, printerSettingsRelId) +
    headerFooterXml(sheet.headerFooter) +
    // Schema order near the tail: <drawing> (the images), then <legacyDrawing> (the VML holding the
    // note boxes), then <tableParts>.
    (drawingRelId !== null ? `<drawing r:id="${drawingRelId}"/>` : '') +
    (legacyDrawingRelId !== null ? `<legacyDrawing r:id="${legacyDrawingRelId}"/>` : '') +
    tablePartsXml(tables) +
    '</worksheet>'
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
  hyperlinks: readonly PlannedHyperlink[]
): string {
  const rels = [
    ...tables.map(({relId, number}) => relationship(relId, REL.table, `../tables/table${number}.xml`)),
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
  // headerRowCount defaults to 1 in OOXML, so only a headerless table needs it stated.
  const headerRowCount = table.headerRow ? '' : ' headerRowCount="0"';
  // totalsRowShown defaults to true; a table without a totals row must say so explicitly.
  const totals = table.totalsRow ? ' totalsRowCount="1"' : ' totalsRowShown="0"';
  const autoFilter =
    table.autoFilterRef !== null ? `<autoFilter ref="${table.autoFilterRef}"/>` : '';
  const columns = table.columns.map((column, i) => tableColumnXml(column, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<table xmlns="${NS.main}" id="${id}" name="${name}" displayName="${name}" ` +
    `ref="${table.ref}"${headerRowCount}${totals}>` +
    autoFilter +
    `<tableColumns count="${table.columns.length}">${columns}</tableColumns>` +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" ' +
    'showRowStripes="1" showColumnStripes="0"/>' +
    '</table>'
  );
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
  const cols: string[] = [];
  for (const {index, properties} of sheet.columns()) {
    // OOXML has no column past XFD (16384); a definition beyond it is corrupt to Excel,
    // so drop it rather than emit an out-of-range <col> range.
    if (index > MAX_COLUMN) continue;
    const col = colXml(index, properties, styles);
    if (col !== '') cols.push(col);
  }
  return cols.length === 0 ? '' : `<cols>${cols.join('')}</cols>`;
}

function colXml(index: number, properties: ColumnProperties, styles: StyleRegistry): string {
  let attrs = `min="${index}" max="${index}"`;
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
  // A <col> with no width, visibility, or style says nothing; omit it entirely.
  return meaningful ? `<col ${attrs}/>` : '';
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

function cellXml(cell: Cell, style: number, shared?: SharedFormulaRole): string {
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
    return `<c r="${ref}"${s}><v>${numberText(value)}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === 'string') {
    return `<c r="${ref}"${s} t="inlineStr"><is>${textElement(value)}</is></c>`;
  }
  if (isRichTextValue(value)) {
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
    cell.protection !== undefined
  );
}

function formulaCellXml(ref: string, s: string, formula: string, result: FormulaResult | undefined): string {
  return formulaBodyXml(ref, s, `<f>${escapeText(mangleFormula(formula))}</f>`, result);
}

// Wrap a prepared `<f>` element (a plain formula, or a shared master/slave `<f>`) with the cell
// element and its cached result, typing the cell by the result's kind exactly as a bare value of that
// kind would be.
function formulaBodyXml(ref: string, s: string, f: string, result: FormulaResult | undefined): string {
  if (result === undefined) {
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
