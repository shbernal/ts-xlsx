// Streaming row reader: yield one worksheet's rows one at a time, without ever building the
// whole {@link Workbook} model.
//
// `readXlsx` materialises every cell of every sheet as a live `Cell` object held in nested Maps —
// fine for editing, but for a large sheet read purely to extract its data it holds the entire grid
// in memory at once. This reader instead *pulls* the sheet's XML through `xmlEvents` and yields a
// plain {@link StreamedRow} at each `</row>`, retaining only the row currently in hand. Peak model
// memory is one row, not the sheet.
//
// Scope of this slice: the package is still inflated whole (bounded by the running counter in
// `./inflate.ts`) and shared strings / styles are read as whole parts — both are legitimately
// document-sized and cheap. What this avoids is retaining N materialised cells. A later slice can
// make the inflate itself per-part lazy; the pull primitive this stands on (`xmlEvents`) is the
// same one that path will use.

import {strFromU8} from 'fflate';

import {decodeAddress} from '../../core/address.ts';
import type {CellValue} from '../../core/value.ts';
import {decodeCellContent} from './cell-value.ts';
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

/** One non-empty cell in a {@link StreamedRow}. */
export interface StreamedCell {
  /** 1-based column index. */
  readonly col: number;
  /** Canonical A1 address (`"B3"`). */
  readonly address: string;
  /** The decoded value — identical to what `readXlsx` would produce for the same cell. */
  readonly value: CellValue;
}

/** One worksheet row, as yielded by {@link readSheetRows}. */
export interface StreamedRow {
  /** 1-based row index. */
  readonly number: number;
  /** The row's non-empty cells, in column order. An empty (or purely style-only) row yields none. */
  readonly cells: readonly StreamedCell[];
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
  const cap = options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  const files = inflatePackage(data, cap);
  const text = (path: string): string | undefined => {
    const bytes = files[path];
    return bytes === undefined ? undefined : strFromU8(bytes);
  };

  const workbookXml = text('xl/workbook.xml');
  if (workbookXml === undefined) throw new Error('not an xlsx package: xl/workbook.xml is missing');

  const chosen = pickSheet(parseWorkbookSheets(workbookXml), options.sheet);
  const rels = parseRelationships(text('xl/_rels/workbook.xml.rels') ?? '');
  const target = rels.get(chosen.relId);
  const path = target === undefined ? undefined : resolveWorkbookPart(target);
  const sheetXml = path === undefined ? undefined : text(path);
  // The sheet is named but its part is missing (a truncated or foreign package) — it has no rows.
  if (sheetXml === undefined) return;

  const sharedStrings = parseSharedStrings(text('xl/sharedStrings.xml') ?? '');
  const xfStyles = parseStyleTable(text('xl/styles.xml') ?? '');
  yield* streamRows(sheetXml, sharedStrings, xfStyles);
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

// Pull the sheet XML through the event stream, yielding a StreamedRow at each `</row>`. The cell
// state mirrors the buffered reader's `parseWorksheet` (same self-closing-`<c/>` handling, same
// capture flags), but commits into a row buffer that is handed off and discarded per row rather
// than into a persistent Worksheet — that hand-off is what bounds retained memory to one row.
function* streamRows(
  xml: string,
  sharedStrings: readonly string[],
  xfStyles: ReadonlyArray<XfStyle>
): Generator<StreamedRow> {
  let rowNumber = 0;
  let lastRow = 0;
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
    const numFmt = cellStyle >= 0 ? xfStyles[cellStyle]?.numFmt : undefined;
    const value = decodeCellContent(
      {type: cellType, hasFormula, formula, hasValue, valueText, inlineText},
      sharedStrings,
      numFmt
    );
    // A blank or purely style-only cell decodes to null; a data read wants only cells that carry
    // something (a formula object, an empty string, a false, and a 0 all count — only null drops).
    if (value !== null) cells.push({col, address: cellRef, value});
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
          cells = [];
          break;
        }
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
        yield {number: rowNumber, cells};
        break;
      default:
        break;
    }
    capture = false;
  }
}
