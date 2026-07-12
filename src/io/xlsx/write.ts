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
import type {Table, TableColumn} from '../../core/table.ts';
import {detectValueType, type FormulaResult, isFormulaValue} from '../../core/value.ts';
import type {SheetProtection, SheetProtectionFlags} from '../../core/protection.ts';
import type {Workbook, WorkbookProperties} from '../../core/workbook.ts';
import type {
  ColumnProperties,
  HeaderFooter,
  PageMargins,
  RowProperties,
  Worksheet,
  WorksheetProperties,
} from '../../core/worksheet.ts';
import {THEME1_XML} from './static-parts.ts';
import {StyleRegistry} from './styles.ts';
import {escapeAttr, escapeText, needsSpacePreserve, XML_DECLARATION} from './xml.ts';

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
} as const;

const REL = {
  worksheet: `${NS.docRels}/worksheet`,
  styles: `${NS.docRels}/styles`,
  theme: `${NS.docRels}/theme`,
  officeDocument: `${NS.docRels}/officeDocument`,
  coreProps: `${NS.packageRels}/metadata/core-properties`,
  extProps: `${NS.docRels}/extended-properties`,
  table: `${NS.docRels}/table`,
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

  // Serialise the worksheets first: interning each cell/row fill into the style table is a
  // side effect of that pass, so styles.xml can only be generated once every sheet is done.
  const styles = new StyleRegistry();
  const sheetXml = sheets.map((sheet, i) => worksheetXml(sheet, sheetTables[i] ?? [], styles));

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypesXml(sheets.length, allTables)),
    '_rels/.rels': strToU8(rootRelsXml()),
    'docProps/core.xml': strToU8(corePropsXml(workbook.properties)),
    'docProps/app.xml': strToU8(appPropsXml()),
    'xl/workbook.xml': strToU8(workbookXml(sheets)),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelsXml(sheets.length)),
    'xl/styles.xml': strToU8(styles.toXml()),
    'xl/theme/theme1.xml': strToU8(THEME1_XML),
  };
  sheets.forEach((_sheet, i) => {
    const tables = sheetTables[i] ?? [];
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml[i] as string);
    if (tables.length > 0) {
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(worksheetRelsXml(tables));
    }
  });
  for (const {table, number} of allTables) {
    files[`xl/tables/table${number}.xml`] = strToU8(tableXml(table, number));
  }

  return zipSync(files, {level: 6});
}

// A table paired with the identifiers the package needs: a workbook-global part number
// and the sheet-local relationship id that links its worksheet to the table part.
interface PlannedTable {
  readonly table: Table;
  readonly number: number;
  readonly relId: string;
}

function contentTypesXml(sheetCount: number, tables: readonly PlannedTable[]): string {
  const overrides = [
    override('/xl/workbook.xml', CT.workbook),
    ...range(sheetCount).map(i => override(`/xl/worksheets/sheet${i + 1}.xml`, CT.worksheet)),
    ...tables.map(({number}) => override(`/xl/tables/table${number}.xml`, CT.table)),
    override('/xl/theme/theme1.xml', CT.theme),
    override('/xl/styles.xml', CT.styles),
    override('/docProps/core.xml', CT.core),
    override('/docProps/app.xml', CT.app),
  ].join('');
  return (
    XML_DECLARATION +
    `<Types xmlns="${NS.contentTypes}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
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

function workbookXml(sheets: readonly Worksheet[]): string {
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
    '</workbook>'
  );
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

function worksheetXml(sheet: Worksheet, tables: readonly PlannedTable[], styles: StyleRegistry): string {
  // A merge overlapping a table is Excel-invalid geometry; reject it before serialising
  // rather than emit a package a consumer repairs on open.
  validateMerges(sheet);

  // Column number formats are defaults a cell inherits unless it overrides them; the writer
  // composes each cell's full style up front so a cell that overrides one facet (say a fill)
  // still carries the column's format, rather than silently dropping it.
  const columnNumFmt = new Map<number, string>();
  for (const {index, properties} of sheet.columns()) {
    if (properties.numFmt !== undefined) columnNumFmt.set(index, properties.numFmt);
  }

  const rowXml: string[] = [];
  let top = Infinity;
  let left = Infinity;
  let bottom = -Infinity;
  let right = -Infinity;

  for (const {number, cells, properties} of sheet.rows()) {
    const populated = cells.filter(cell => cell.value !== null);
    const attrs = rowAttrs(properties, styles);
    // A row with neither data nor its own formatting has nothing to serialise.
    if (populated.length === 0 && attrs === '') continue;
    const rowFill = properties?.fill;
    const cellsXml = populated
      .map(cell => {
        // Cell overrides win over row/column defaults; a cell with any facet gets its own,
        // fully-composed style entry so no default facet is lost to the override.
        const style = styles.styleId({
          fill: cell.fill ?? rowFill,
          numFmt: cell.numFmt ?? columnNumFmt.get(cell.col),
          font: cell.font,
          border: cell.border,
          alignment: cell.alignment,
          protection: cell.protection,
        });
        return cellXml(cell, style);
      })
      .join('');
    rowXml.push(`<row r="${number}"${attrs}>${cellsXml}</row>`);
    for (const cell of populated) {
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
    `<dimension ref="${dimensionRef}"/>` +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    sheetFormatPr(sheet.properties) +
    colsXml(sheet, styles) +
    sheetData +
    sheetProtectionXml(sheet.protection) +
    mergeCellsXml(sheet.merges) +
    pageMarginsXml(sheet.pageMargins) +
    headerFooterXml(sheet.headerFooter) +
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

function mergeCellsXml(merges: readonly string[]): string {
  if (merges.length === 0) return '';
  const cells = merges.map(range => `<mergeCell ref="${escapeAttr(decodeRange(range).dimensions)}"/>`).join('');
  return `<mergeCells count="${merges.length}">${cells}</mergeCells>`;
}

// Each sheet-protection flag maps to a `<sheetProtection>` attribute whose value is INVERTED
// from the author-facing allow-flag: the attribute records that an operation is *forbidden*
// ("1"), so `allow: true` serialises as "0". Only a value that differs from OOXML's per-
// attribute default is written — most editing operations default to forbidden under
// protection (`defaultForbidden: true`), while selecting cells defaults to permitted.
const PROTECTION_FLAGS: readonly {readonly key: keyof SheetProtectionFlags; readonly defaultForbidden: boolean}[] = [
  {key: 'formatCells', defaultForbidden: true},
  {key: 'formatColumns', defaultForbidden: true},
  {key: 'formatRows', defaultForbidden: true},
  {key: 'insertColumns', defaultForbidden: true},
  {key: 'insertRows', defaultForbidden: true},
  {key: 'insertHyperlinks', defaultForbidden: true},
  {key: 'deleteColumns', defaultForbidden: true},
  {key: 'deleteRows', defaultForbidden: true},
  {key: 'sort', defaultForbidden: true},
  {key: 'autoFilter', defaultForbidden: true},
  {key: 'pivotTables', defaultForbidden: true},
  {key: 'objects', defaultForbidden: false},
  {key: 'scenarios', defaultForbidden: false},
  {key: 'selectLockedCells', defaultForbidden: false},
  {key: 'selectUnlockedCells', defaultForbidden: false},
];

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
  for (const {key, defaultForbidden} of PROTECTION_FLAGS) {
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

function worksheetRelsXml(tables: readonly PlannedTable[]): string {
  const rels = tables
    .map(({relId, number}) => relationship(relId, REL.table, `../tables/table${number}.xml`))
    .join('');
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
  // A column-level number format is carried as the column's own style; its cells inherit it
  // via the composition above, and Excel applies it to the column's empty cells too.
  const style = styles.styleId({numFmt: properties.numFmt});
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

function cellXml(cell: Cell, style: number): string {
  const ref = cell.address;
  const value = cell.value;
  const s = style !== 0 ? ` s="${style}"` : '';

  if (isFormulaValue(value)) {
    return formulaCellXml(ref, s, value.formula, value.result);
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
  throw new Error(`writing a ${detectValueType(value)} cell value is not implemented yet`);
}

function formulaCellXml(ref: string, s: string, formula: string, result: FormulaResult | undefined): string {
  const f = `<f>${escapeText(formula)}</f>`;
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
  throw new Error('writing a non-primitive formula result is not implemented yet');
}

function textElement(value: string): string {
  const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : '';
  return `<t${space}>${escapeText(value)}</t>`;
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
