// Workbook-level serialisation: the package's `[Content_Types].xml`, its root and workbook `.rels`
// parts, `xl/workbook.xml` (sheets, defined names, calc/protection settings, pivot-cache and slicer
// registrations), and the `docProps` core/app property parts.

import {mangleFormula} from '../../core/formula.ts';
import type {Workbook, WorkbookProperties} from '../../core/workbook.ts';
import {WORKBOOK_PROTECTION_CREDENTIAL_ATTRS} from '../../core/workbook-protection.ts';
import {imageContentType} from './images.ts';
import {SLICER_CACHES_EXT_URI, X14_NS} from './namespaces.ts';
import type {
  PivotPlan,
  PlannedTable,
  PreservedPartPlan,
  PreservedWorkbookReferencePlan,
} from './package-plan.ts';
import {range, relativePartPath} from './part-paths.ts';
import {NS, REL, relationship} from './relationships.ts';
import {escapeAttr, escapeText, XML_DECLARATION} from './xml.ts';

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

// A preserved workbook reference with the relationship id assigned for emission (see the body and
// rels-part wiring in `buildPackageParts`).
export type PreservedWorkbookRel = PreservedWorkbookReferencePlan & {readonly relId: string};

export function contentTypesXml(
  sheetCount: number,
  tables: readonly PlannedTable[],
  commentNumbers: readonly number[],
  drawingNumbers: readonly number[],
  printerSettingsNumbers: readonly number[],
  mediaExtensions: readonly string[],
  hasSharedStrings: boolean,
  preservedParts: readonly PreservedPartPlan[],
  pivots: readonly PivotPlan[],
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
    ...range(sheetCount).map((i) => override(`/xl/worksheets/sheet${i + 1}.xml`, CT.worksheet)),
    ...tables.map(({number}) => override(`/xl/tables/table${number}.xml`, CT.table)),
    ...drawingNumbers.map((number) => override(`/xl/drawings/drawing${number}.xml`, CT.drawing)),
    ...commentNumbers.map((number) => override(`/xl/comments${number}.xml`, CT.comments)),
    ...pivots.map(({number}) => override(`/xl/pivotTables/pivotTable${number}.xml`, CT.pivotTable)),
    ...pivots.map(({number}) =>
      override(`/xl/pivotCache/pivotCacheDefinition${number}.xml`, CT.pivotCacheDefinition),
    ),
    ...pivots.map(({number}) =>
      override(`/xl/pivotCache/pivotCacheRecords${number}.xml`, CT.pivotCacheRecords),
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
    printerSettingsNumbers.length > 0
      ? `<Default Extension="bin" ContentType="${CT.printerSettings}"/>`
      : '';
  const imageDefaults = mediaExtensions
    .map((ext) => `<Default Extension="${ext}" ContentType="${imageContentType(ext)}"/>`)
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

export function rootRelsXml(): string {
  const rels = [
    relationship('rId1', REL.officeDocument, 'xl/workbook.xml'),
    relationship('rId2', REL.coreProps, 'docProps/core.xml'),
    relationship('rId3', REL.extProps, 'docProps/app.xml'),
  ].join('');
  return `${XML_DECLARATION}<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

export function workbookXml(
  workbook: Workbook,
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[],
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
    workbookProtectionXml(workbook) +
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
  const caches = preservedRels.filter((ref) => ref.relType.endsWith('/slicerCache'));
  if (caches.length === 0) return '';
  const entries = caches.map((ref) => `<x14:slicerCache r:id="${ref.relId}"/>`).join('');
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
  pivots: readonly PivotPlan[],
): string {
  const preserved = preservedRels
    .filter((ref) => ref.pivotCacheId !== undefined)
    .map(
      (ref) =>
        `<pivotCache cacheId="${escapeAttr(ref.pivotCacheId as string)}" r:id="${ref.relId}"/>`,
    );
  const generated = pivots.map(
    (pivot) => `<pivotCache cacheId="${escapeAttr(pivot.cacheId)}" r:id="${pivot.workbookRelId}"/>`,
  );
  const entries = [...preserved, ...generated];
  if (entries.length === 0) return '';
  return `<pivotCaches>${entries.join('')}</pivotCaches>`;
}

// `<workbookProtection>` precedes `<sheets>` in CT_Workbook order. It re-emits the workbook's
// structure/window lock flags (each written only when true, so an unlocked aspect stays absent) and
// the preserved password/agile-hash credential attributes verbatim. Emitted only when the workbook
// actually declares protection — the flags or a credential — so an unprotected workbook stays clean.
function workbookProtectionXml(workbook: Workbook): string {
  const p = workbook.protection;
  if (p === undefined) return '';
  const attrs: string[] = [];
  if (p.lockStructure) attrs.push('lockStructure="1"');
  if (p.lockWindows) attrs.push('lockWindows="1"');
  if (p.lockRevision) attrs.push('lockRevision="1"');
  for (const key of WORKBOOK_PROTECTION_CREDENTIAL_ATTRS) {
    const value = p.credentials?.[key];
    if (value !== undefined) attrs.push(`${key}="${escapeAttr(value)}"`);
  }
  if (attrs.length === 0) return '';
  return `<workbookProtection ${attrs.join(' ')}/>`;
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
  const userEntries = workbook.definedNames.map((name) => {
    const scopeAttr =
      name.scope === undefined
        ? ''
        : ` localSheetId="${sheets.findIndex((sheet) => sheet.name === name.scope)}"`;
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
        ],
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

export function workbookRelsXml(
  sheetCount: number,
  hasSharedStrings: boolean,
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[],
): string {
  const rels = [
    ...range(sheetCount).map((i) =>
      relationship(`rId${i + 1}`, REL.worksheet, `worksheets/sheet${i + 1}.xml`),
    ),
    relationship(`rId${sheetCount + 1}`, REL.styles, 'styles.xml'),
    relationship(`rId${sheetCount + 2}`, REL.theme, 'theme/theme1.xml'),
    ...(hasSharedStrings
      ? [relationship(`rId${sheetCount + 3}`, REL.sharedStrings, 'sharedStrings.xml')]
      : []),
    // A preserved cache's target is package-absolute; express it relative to the workbook part.
    ...preservedRels.map((ref) =>
      relationship(
        ref.relId,
        ref.relType,
        escapeAttr(relativePartPath('xl/workbook.xml', ref.entryPath)),
      ),
    ),
    // A generated pivot cache's workbook relationship reaches its cache definition part.
    ...pivots.map((pivot) =>
      relationship(
        pivot.workbookRelId,
        REL.pivotCacheDefinition,
        `pivotCache/pivotCacheDefinition${pivot.number}.xml`,
      ),
    ),
  ].join('');
  return `${XML_DECLARATION}<Relationships xmlns="${NS.packageRels}">${rels}</Relationships>`;
}

export function corePropsXml(properties: WorkbookProperties): string {
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

export function appPropsXml(): string {
  return (
    XML_DECLARATION +
    `<Properties xmlns="${NS.extProps}"><Application>ts-xlsx</Application></Properties>`
  );
}
