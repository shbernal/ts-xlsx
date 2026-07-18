// Streaming row reader: yield a worksheet's rows one at a time, without ever building the whole
// {@link Workbook} model.
//
// `readXlsx` materialises every cell of every sheet as a live `Cell` object held in nested Maps —
// fine for editing, but for a large sheet read purely to extract its data it holds the entire grid
// in memory at once. This reader instead *pulls* the sheet's XML through `xmlEvents` and yields a
// plain {@link StreamedRow} at each `</row>`, retaining only the row currently in hand. Peak model
// memory is one row, not the sheet.
//
// Two entry points sit on the same scanner:
//   - {@link readSheetRows} streams a single selected sheet's rows (the terse data-extraction case).
//   - {@link readWorkbookStream} yields a {@link StreamedSheet} per worksheet in workbook order, so a
//     caller can walk every sheet — each sheet's rows still stream one at a time.
//
// Scope of this slice: the package is still inflated whole (bounded by the running counter in
// `./inflate.ts`) and shared strings / styles are read as whole parts — both are legitimately
// document-sized and cheap. What this avoids is retaining N materialised cells. A later slice can
// make the inflate itself per-part lazy; the pull primitive this stands on (`xmlEvents`) is the
// same one that path will use.

import {strFromU8} from 'fflate';

import {decodeAddress, MAX_COLUMN} from '../../core/address.ts';
import type {CellValue} from '../../core/value.ts';
import {decodeCellContent, type SharedString} from './cell-value.ts';
import {inflatePackage} from './inflate.ts';
import {
  DEFAULT_MAX_UNCOMPRESSED,
  parseRelationships,
  parseSharedStrings,
  parseStyleTable,
  parseWorkbookSheets,
  type ReadXlsxOptions,
  resolveWorkbookPart,
  type XfStyle,
} from './read.ts';
import {localName, xmlEvents} from './xml-read.ts';

export interface ReadSheetRowsOptions extends ReadXlsxOptions {
  /**
   * Which worksheet to stream: its name, or its 1-based position in the workbook. Defaults to the
   * first sheet.
   */
  readonly sheet?: string | number;
}

/**
 * The resolved style facets of a streamed cell — its own `<c s>` cell format, flattened exactly as
 * the buffered reader resolves it. Present only when the cell carries a format; a consumer can copy
 * these straight onto a writer cell to preserve its look through a streaming read→write.
 */
export type StreamedCellStyle = XfStyle;

/** One non-empty cell in a {@link StreamedRow}. */
export interface StreamedCell {
  /** 1-based column index. */
  readonly col: number;
  /** Canonical A1 address (`"B3"`). */
  readonly address: string;
  /** The decoded value — identical to what `readXlsx` would produce for the same cell. */
  readonly value: CellValue;
  /** The cell's resolved style facets, or absent when the cell carries no format of its own. */
  readonly style?: StreamedCellStyle;
}

/** One worksheet row, as yielded by {@link readSheetRows} / {@link StreamedSheet.rows}. */
export interface StreamedRow {
  /** 1-based row index. */
  readonly number: number;
  /** Whether the row declares itself hidden. */
  readonly hidden: boolean;
  /** The row's non-empty cells, in column order. An empty (or purely style-only) row yields none. */
  readonly cells: readonly StreamedCell[];
}

/**
 * One worksheet, as yielded by {@link readWorkbookStream}. The sheet's {@link rows} stream one at a
 * time; its {@link hiddenColumns} and {@link merges} are populated by that same single pass.
 *
 * The two summaries are resolved lazily: reading either accessor drives a full scan of the sheet if
 * its rows have not already been consumed, so their order relative to `rows()` never matters. (When
 * rows *are* consumed first — the streaming idiom — the accessors reuse that pass and re-scan
 * nothing.)
 */
export interface StreamedSheet {
  /** The worksheet's declared name, joined from the workbook part — never a positional placeholder. */
  readonly name: string;
  /** Stream this sheet's rows, one at a time, in sheet order. */
  rows(): Generator<StreamedRow>;
  /** 1-based indices of columns the sheet declares hidden, ascending. */
  readonly hiddenColumns: readonly number[];
  /** The sheet's merged ranges, as canonical A1 range strings, in declaration order. */
  readonly merges: readonly string[];
}

/**
 * Stream a worksheet's rows from an `.xlsx` package, yielding each in sheet order without building
 * the workbook model. Only rows the sheet actually declares are yielded, and within a row only its
 * non-empty cells — a blank or style-only cell contributes nothing, matching the intent of a data
 * read.
 *
 * @param data The raw `.xlsx` bytes.
 * @param options Sheet selector and the inflate bound (see {@link ReadSheetRowsOptions}).
 * @throws {Error} if the archive is malformed, exceeds the inflate bound, or names no worksheet —
 *   or if `options.sheet` selects a sheet that does not exist.
 */
export function* readSheetRows(
  data: Uint8Array,
  options: ReadSheetRowsOptions = {}
): Generator<StreamedRow> {
  const pkg = openPackage(data, options.maxUncompressedBytes);
  const chosen = pickSheet(pkg.sheets, options.sheet);
  const sheetXml = pkg.sheetXml(chosen.relId);
  // The sheet is named but its part is missing (a truncated or foreign package) — it has no rows.
  if (sheetXml === undefined) return;
  yield* scanSheet(sheetXml, pkg.sharedStrings, pkg.xfStyles, new Set(), []);
}

/**
 * Stream every worksheet of an `.xlsx` package in workbook order, without building the workbook
 * model. Each yielded {@link StreamedSheet} carries the declared sheet name and lets the caller
 * stream that sheet's rows and read its hidden-column and merge summaries — the streaming analogue
 * of walking `readXlsx(data).worksheets`.
 *
 * @param data The raw `.xlsx` bytes.
 * @param options The inflate bound (see {@link ReadXlsxOptions}).
 * @throws {Error} if the archive is malformed or exceeds the inflate bound.
 */
export function* readWorkbookStream(
  data: Uint8Array,
  options: ReadXlsxOptions = {}
): Generator<StreamedSheet> {
  const pkg = openPackage(data, options.maxUncompressedBytes);
  for (const sheet of pkg.sheets) {
    // A named sheet whose part is missing (truncated/foreign package) still surfaces — with no rows,
    // no hidden columns, and no merges — rather than vanishing from the workbook's sheet list.
    const xml = pkg.sheetXml(sheet.relId) ?? '';
    yield new StreamedSheetReader(sheet.name, xml, pkg.sharedStrings, pkg.xfStyles);
  }
}

// The shared parts every streaming read needs: the sheet directory (name + rel id, in workbook
// order), the shared-string and style tables, and a resolver from a sheet's rel id to its XML. The
// package is inflated once; sheet XML is fetched lazily so a sheet the caller never visits is never
// stringified.
interface OpenPackage {
  readonly sheets: ReadonlyArray<{name: string; relId: string}>;
  readonly sharedStrings: readonly SharedString[];
  readonly xfStyles: ReadonlyArray<XfStyle>;
  sheetXml(relId: string): string | undefined;
}

function openPackage(data: Uint8Array, maxUncompressedBytes: number | undefined): OpenPackage {
  const cap = maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  const files = inflatePackage(data, cap);
  const text = (path: string): string | undefined => {
    const bytes = files[path];
    return bytes === undefined ? undefined : strFromU8(bytes);
  };

  const workbookXml = text('xl/workbook.xml');
  if (workbookXml === undefined) throw new Error('not an xlsx package: xl/workbook.xml is missing');

  const sheets = parseWorkbookSheets(workbookXml);
  const rels = parseRelationships(text('xl/_rels/workbook.xml.rels') ?? '');
  const sharedStrings = parseSharedStrings(text('xl/sharedStrings.xml') ?? '');
  const {cellXfs: xfStyles} = parseStyleTable(text('xl/styles.xml') ?? '');

  return {
    sheets,
    sharedStrings,
    xfStyles,
    sheetXml(relId: string): string | undefined {
      const target = rels.get(relId);
      const path = target === undefined ? undefined : resolveWorkbookPart(target);
      return path === undefined ? undefined : text(path);
    },
  };
}

function pickSheet(
  sheets: ReadonlyArray<{name: string; relId: string}>,
  selector: string | number | undefined
): {name: string; relId: string} {
  if (sheets.length === 0) throw new Error('workbook names no worksheets');
  if (selector === undefined) return sheets[0] as {name: string; relId: string};
  if (typeof selector === 'number') {
    const sheet = sheets[selector - 1];
    if (sheet === undefined) throw new RangeError(`no worksheet at position ${selector}`);
    return sheet;
  }
  const sheet = sheets.find((candidate) => candidate.name === selector);
  if (sheet === undefined) throw new Error(`no worksheet named ${JSON.stringify(selector)}`);
  return sheet;
}

// A single worksheet exposed by readWorkbookStream. Its rows() re-scans on each call (a fresh pass,
// so it is safely re-iterable); the hidden-column and merge accessors reuse a completed scan or, if
// the rows were never drained, drive one of their own. The hidden/merge state is filled in by the
// same scanSheet pass that yields the rows.
class StreamedSheetReader implements StreamedSheet {
  readonly name: string;
  readonly #xml: string;
  readonly #sharedStrings: readonly SharedString[];
  readonly #xfStyles: ReadonlyArray<XfStyle>;
  #hiddenColumns = new Set<number>();
  #merges: string[] = [];
  #scanned = false;

  constructor(
    name: string,
    xml: string,
    sharedStrings: readonly SharedString[],
    xfStyles: ReadonlyArray<XfStyle>
  ) {
    this.name = name;
    this.#xml = xml;
    this.#sharedStrings = sharedStrings;
    this.#xfStyles = xfStyles;
  }

  *rows(): Generator<StreamedRow> {
    this.#hiddenColumns = new Set();
    this.#merges = [];
    this.#scanned = false;
    yield* scanSheet(this.#xml, this.#sharedStrings, this.#xfStyles, this.#hiddenColumns, this.#merges);
    this.#scanned = true;
  }

  get hiddenColumns(): readonly number[] {
    this.#ensureScanned();
    return [...this.#hiddenColumns].sort((a, b) => a - b);
  }

  get merges(): readonly string[] {
    this.#ensureScanned();
    return this.#merges;
  }

  // Drain a scan purely for its summaries when the caller reads them without (or before) iterating
  // rows. A completed row iteration already set #scanned, so this re-scans nothing in the common
  // streaming idiom.
  #ensureScanned(): void {
    if (this.#scanned) return;
    for (const _row of this.rows()) {
      // The rows themselves are irrelevant here — we only want the hidden/merge side effects.
    }
  }
}

// Pull the sheet XML through the event stream, yielding a StreamedRow at each `</row>`, while
// recording the sheet's hidden columns (from `<col hidden>`, before <sheetData>) and merged ranges
// (from `<mergeCells>`, after <sheetData>) into the caller-supplied collectors. The cell state
// mirrors the buffered reader's `parseWorksheet` (same self-closing-`<c/>` handling, same capture
// flags), but commits into a row buffer that is handed off and discarded per row rather than into a
// persistent Worksheet — that hand-off is what bounds retained memory to one row.
function* scanSheet(
  xml: string,
  sharedStrings: readonly SharedString[],
  xfStyles: ReadonlyArray<XfStyle>,
  hiddenColumns: Set<number>,
  merges: string[]
): Generator<StreamedRow> {
  let rowNumber = 0;
  let lastRow = 0;
  let rowHidden = false;
  let cells: StreamedCell[] = [];

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

  const finalizeCell = (): void => {
    if (cellRef === '') return;
    const {col} = decodeAddress(cellRef);
    if (col === undefined) return;
    const style = cellStyle >= 0 ? xfStyles[cellStyle] : undefined;
    const value = decodeCellContent(
      {type: cellType, hasFormula, formula, hasValue, valueText, inlineText},
      sharedStrings,
      style?.numFmt
    );
    // A blank or purely style-only cell decodes to null; a data read wants only cells that carry
    // something (a formula object, an empty string, a false, and a 0 all count — only null drops).
    if (value !== null) {
      cells.push(style ? {col, address: cellRef, value, style} : {col, address: cellRef, value});
    }
  };

  for (const event of xmlEvents(xml)) {
    if (event.kind === 'text') {
      if (capture) text += event.text;
      continue;
    }
    if (event.kind === 'open') {
      const local = localName(event.name);
      text = '';
      capture = false;
      switch (local) {
        case 'row': {
          const declared = Number(event.attrs.r);
          rowNumber = Number.isInteger(declared) && declared >= 1 ? declared : lastRow + 1;
          lastRow = rowNumber;
          rowHidden = event.attrs.hidden === '1' || event.attrs.hidden === 'true';
          cells = [];
          break;
        }
        case 'col':
          collectHiddenColumn(event.attrs, hiddenColumns);
          break;
        case 'mergeCell':
          if (event.attrs.ref !== undefined) merges.push(event.attrs.ref);
          break;
        case 'c':
          cellRef = event.attrs.r ?? '';
          cellType = event.attrs.t ?? '';
          cellStyle = event.attrs.s !== undefined ? Number(event.attrs.s) : -1;
          formula = '';
          valueText = '';
          inlineText = '';
          hasFormula = false;
          hasValue = false;
          // A self-closing `<c .../>` fires no close, so commit it here (mirrors parseWorksheet).
          if (event.selfClosing) finalizeCell();
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
        default:
          break;
      }
      if (event.selfClosing && (local === 'f' || local === 'v')) capture = false;
      continue;
    }
    // close
    const local = localName(event.name);
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
      case 'c':
        finalizeCell();
        break;
      case 'row':
        yield {number: rowNumber, hidden: rowHidden, cells};
        break;
      default:
        break;
    }
    capture = false;
  }
}

// Record the hidden columns a `<col min max hidden>` element declares. The span is clamped to the
// format's column ceiling and gathered into a Set, so even a hostile file full of full-width hidden
// spans can add at most MAX_COLUMN distinct entries — never an unbounded allocation.
function collectHiddenColumn(
  attrs: {readonly [k: string]: string | undefined},
  hiddenColumns: Set<number>
): void {
  if (attrs.hidden !== '1' && attrs.hidden !== 'true') return;
  const min = Number(attrs.min);
  const max = Number(attrs.max);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1) return;
  const last = Math.min(max, MAX_COLUMN);
  for (let index = min; index <= last; index++) hiddenColumns.add(index);
}
