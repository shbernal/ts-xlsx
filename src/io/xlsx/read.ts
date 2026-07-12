// The buffered `.xlsx` reader: an OPC zip package in, a Workbook model out.
//
// It reconstructs the part of the model the writer emits today — sheet names and order,
// cells holding a number, string, boolean, or formula, per-column width/visibility,
// per-row height/visibility, merged ranges, page margins, and pattern-fill styles (per
// cell, or inherited from a formatted row). Fonts/borders/number formats, shared-formula
// slaves, and the richer value kinds land as the model grows; an unrecognised construct
// is skipped rather than guessed, so a foreign file reads without crashing even where a
// facet is not yet materialised.
//
// Untrusted input: inflate is bounded here (a declared-size cap rejects the naïve zip
// bomb), and the parser (ADR 0004) never expands entities. See ADR 0004 for the honest
// limits of the declared-size bound.

import {unzipSync, strFromU8, type UnzipFileInfo} from 'fflate';

import {decodeAddress} from '../../core/address.ts';
import type {Color, Fill, FillPatternType} from '../../core/style.ts';
import {type CellValue, type FormulaResult, isErrorCode} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import type {PageMargins, Worksheet} from '../../core/worksheet.ts';
import {localName, parseXml} from './xml-read.ts';

export interface ReadXlsxOptions {
  /**
   * Maximum total *declared* uncompressed bytes across all package parts. A zip that
   * claims to inflate past this is rejected as a probable bomb before the parser runs.
   * Defaults to 512 MiB.
   */
  readonly maxUncompressedBytes?: number;
}

const DEFAULT_MAX_UNCOMPRESSED = 512 * 1024 * 1024;
const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

/**
 * Read an `.xlsx` package into a {@link Workbook}.
 *
 * @throws {Error} if the archive is malformed, exceeds the inflate bound, or names no
 *   worksheet parts (a workbook with no sheets is not a valid package).
 */
export function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook {
  const cap = options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  let declared = 0;
  const files = unzipSync(data, {
    filter(file: UnzipFileInfo): boolean {
      declared += file.originalSize;
      if (declared > cap) {
        throw new Error(`refusing to inflate: declared uncompressed size exceeds ${cap} bytes (possible zip bomb)`);
      }
      return true;
    },
  });

  const partText = (path: string): string | undefined => {
    const bytes = files[path];
    return bytes === undefined ? undefined : strFromU8(bytes);
  };

  const workbookXml = partText('xl/workbook.xml');
  if (workbookXml === undefined) throw new Error('not an xlsx package: xl/workbook.xml is missing');

  const rels = parseRelationships(partText('xl/_rels/workbook.xml.rels') ?? '');
  const sharedStrings = parseSharedStrings(partText('xl/sharedStrings.xml') ?? '');
  // The style table resolves a cell/row `s` index to its fill; a package without one
  // (a hand-rolled foreign file) yields an empty table and every `s` reads as unstyled.
  const xfFills = parseStyleFills(partText('xl/styles.xml') ?? '');

  const workbook = new Workbook();
  const core = partText('docProps/core.xml');
  if (core !== undefined) applyCoreProperties(workbook, core);

  for (const {name, relId} of parseWorkbookSheets(workbookXml)) {
    const target = rels.get(relId);
    const sheet = workbook.addWorksheet(name);
    const path = target === undefined ? undefined : resolveWorkbookPart(target);
    const sheetXml = path === undefined ? undefined : partText(path);
    if (sheetXml !== undefined) parseWorksheet(sheetXml, sheet, sharedStrings, xfFills);
  }
  return workbook;
}

// A workbook relationship target is relative to the `xl/` directory (`worksheets/sheet1.xml`)
// or absolute from the package root (`/xl/worksheets/sheet1.xml`); normalise both to a part path.
function resolveWorkbookPart(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  return `xl/${target.replace(/^\.\//, '')}`;
}

function parseRelationships(xml: string): Map<string, string> {
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

function parseWorkbookSheets(xml: string): Array<{name: string; relId: string}> {
  const sheets: Array<{name: string; relId: string}> = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) === 'sheet') {
        sheets.push({name: attrs.name ?? '', relId: attrs['r:id'] ?? ''});
      }
    },
    onText() {},
    onClose() {},
  });
  return sheets;
}

// Shared strings resolve `t="s"` cells: each <si> is one entry, its text the concatenation
// of its <t> runs (a plain <si><t>…</t> or a rich <si><r><t>…</t></r>…).
function parseSharedStrings(xml: string): string[] {
  if (xml === '') return [];
  const strings: string[] = [];
  let current = '';
  let capture = false;
  let text = '';
  parseXml(xml, {
    onOpen(name) {
      const local = localName(name);
      if (local === 'si') current = '';
      else if (local === 't') {
        capture = true;
        text = '';
      }
    },
    onText(chunk) {
      if (capture) text += chunk;
    },
    onClose(name) {
      const local = localName(name);
      if (local === 't') {
        current += text;
        capture = false;
      } else if (local === 'si') {
        strings.push(current);
      }
    },
  });
  return strings;
}

// styles.xml is a shared table: <fills> lists the fills, <cellXfs> lists the cell formats,
// each naming a fill by id. We flatten that indirection into one array — cellXfs index →
// resolved fill — so a cell's `s` attribute maps straight to its fill. Only fills are read
// today; other facets on an xf are ignored until their model lands. `<fills>` always
// precedes `<cellXfs>` in the schema, so the fill list is complete before an xf references it.
function parseStyleFills(xml: string): ReadonlyArray<Fill | undefined> {
  if (xml === '') return [];
  const fills: Array<Fill | undefined> = [];
  const xfFills: Array<Fill | undefined> = [];
  let inFills = false;
  let inCellXfs = false;
  let pattern = '';
  let fgColor: Color | undefined;
  let bgColor: Color | undefined;

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      switch (localName(name)) {
        case 'fills':
          inFills = true;
          break;
        case 'cellXfs':
          inCellXfs = true;
          break;
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
        case 'xf':
          // cellStyleXfs also holds <xf>; only those inside <cellXfs> are cell references.
          if (inCellXfs) {
            const fillId = Number(attrs.fillId);
            xfFills.push(Number.isInteger(fillId) ? fills[fillId] : undefined);
          }
          break;
        default:
          break;
      }
    },
    onText() {},
    onClose(name) {
      switch (localName(name)) {
        case 'fills':
          inFills = false;
          break;
        case 'cellXfs':
          inCellXfs = false;
          break;
        case 'patternFill':
          if (inFills) fills.push(toFill(pattern, fgColor, bgColor));
          break;
        default:
          break;
      }
    },
  });
  return xfFills;
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
  sharedStrings: readonly string[],
  xfFills: ReadonlyArray<Fill | undefined>
): void {
  let cellRef = '';
  let cellType = '';
  let cellStyle = -1;
  let formula = '';
  let valueText = '';
  let inlineText = '';
  let hasFormula = false;
  let hasValue = false;
  let inInlineString = false;
  let capture = false;
  let text = '';
  // A row with customFormat="1" supplies a default fill for its cells that carry no `s`.
  let rowStyle = -1;
  let rowCustomFormat = false;

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      const local = localName(name);
      text = '';
      capture = false;
      switch (local) {
        case 'col':
          applyColumn(sheet, attrs);
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
          formula = '';
          valueText = '';
          inlineText = '';
          hasFormula = false;
          hasValue = false;
          break;
        case 'is':
          inInlineString = true;
          inlineText = '';
          break;
        case 'f':
        case 'v':
        case 't':
          capture = true;
          break;
        case 'mergeCell':
          if (attrs.ref !== undefined && attrs.ref !== '') sheet.mergeCells(attrs.ref);
          break;
        case 'pageMargins':
          applyMargins(sheet.pageMargins, attrs);
          break;
        default:
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
          if (inInlineString) inlineText += text;
          break;
        case 'is':
          inInlineString = false;
          break;
        case 'c': {
          if (cellRef === '') break;
          // A cell's own style wins; a cell without one inherits its row's format fill.
          const fill =
            cellStyle >= 0
              ? xfFills[cellStyle]
              : rowCustomFormat && rowStyle >= 0
                ? xfFills[rowStyle]
                : undefined;
          finalizeCell(sheet, cellRef, cellType, hasFormula, formula, hasValue, valueText, inlineText, sharedStrings, fill);
          break;
        }
        case 'row':
          rowStyle = -1;
          rowCustomFormat = false;
          break;
        default:
          break;
      }
      capture = false;
    },
  });
}

function applyColumn(sheet: Worksheet, attrs: {readonly [k: string]: string}): void {
  const min = Number(attrs.min);
  const max = Number(attrs.max);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1) return;
  const width = attrs.width !== undefined ? Number(attrs.width) : undefined;
  const hidden = attrs.hidden === '1' || attrs.hidden === 'true';
  for (let index = min; index <= max; index++) {
    const properties = sheet.getColumn(index);
    if (width !== undefined && Number.isFinite(width) && attrs.customWidth !== '0') properties.width = width;
    if (hidden) properties.hidden = true;
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

function finalizeCell(
  sheet: Worksheet,
  ref: string,
  type: string,
  hasFormula: boolean,
  formula: string,
  hasValue: boolean,
  valueText: string,
  inlineText: string,
  sharedStrings: readonly string[],
  fill: Fill | undefined
): void {
  const {col, row} = decodeAddress(ref);
  if (col === undefined || row === undefined) return;
  const cell = sheet.getCell(ref);
  if (fill !== undefined) cell.fill = fill;

  if (hasFormula) {
    const result = hasValue ? decodeResult(type, valueText) : undefined;
    cell.value = result === undefined ? {formula} : {formula, result};
    return;
  }
  cell.value = decodeValue(type, valueText, inlineText, hasValue, sharedStrings);
}

function decodeValue(
  type: string,
  valueText: string,
  inlineText: string,
  hasValue: boolean,
  sharedStrings: readonly string[]
): CellValue {
  switch (type) {
    case 'inlineStr':
      return inlineText;
    case 'str':
      return valueText;
    case 's': {
      const index = Number(valueText);
      return Number.isInteger(index) ? sharedStrings[index] ?? '' : '';
    }
    case 'b':
      return valueText === '1' || valueText === 'true';
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return hasValue ? Number(valueText) : null;
  }
}

function decodeResult(type: string, valueText: string): FormulaResult {
  switch (type) {
    case 'str':
      return valueText;
    case 'b':
      return valueText === '1' || valueText === 'true';
    case 'e':
      return isErrorCode(valueText) ? {error: valueText} : valueText;
    default:
      return Number(valueText);
  }
}
