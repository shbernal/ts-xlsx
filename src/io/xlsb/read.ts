// The buffered `.xlsb` reader: a binary OPC package in, a {@link Workbook} model out.
//
// An `.xlsb` is the *same* OPC/ZIP container and the *same* relationship graph as an `.xlsx` — the
// workbook part points at its sheets, its shared strings, and its style sheet through ordinary
// `.rels` XML. Only the office-document parts differ: `xl/workbook.bin`, `xl/worksheets/sheetN.bin`
// and friends are BIFF12 record streams instead of XML. So this module reuses the container layer
// wholesale (`../xlsx/read-opc.ts`, the bounded inflater) and swaps only the part parsers, which is
// exactly the "two codecs over one model" shape the format note argues for.
//
// The model produced is the one `readXlsx` produces, not a parallel one: the same `Workbook`, the
// same `XfStyle` table, the same cells. That is what lets a caller convert between the two forms —
// and what the corpus asserts, by reading a workbook Excel saved in both forms and comparing.
//
// Not yet decoded (each its own slice of work, none silently wrong): formula token streams and the
// defined names that are built from them, rich-text runs, tables, pivots, and conditional formatting.

import {Workbook} from '../../core/workbook.ts';
import type {WorksheetState} from '../../core/worksheet.ts';
import {UnsupportedFormatError} from '../xlsx/errors.ts';
import {packageAccessors, parseRelationships, resolveWorkbookPart} from '../xlsx/read-opc.ts';
import {DEFAULT_MAX_UNCOMPRESSED, type ReadXlsxOptions} from '../xlsx/read-options.ts';
import {inflateSpreadsheetPackage} from '../xlsx/sniff-format.ts';
import {RecordReader} from './primitives.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseStyleTable} from './read-styles.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

/** The office-document part every `.xlsb` package is entered through. */
export const XLSB_WORKBOOK_PART = 'xl/workbook.bin';

/**
 * Read an `.xlsb` (binary BIFF12) package into a {@link Workbook}.
 *
 * @throws {UnsupportedFormatError} if the input is not an `.xlsb` package — a legacy `.xls`
 *   (`.format === 'xls'`), an XML `.xlsx` or unrecognised blob (`'unknown'`).
 * @throws {XlsbParseError} if a binary part is malformed.
 * @throws {Error} if the archive exceeds the inflate bound (a probable zip bomb).
 */
export function readXlsb(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook {
  const cap = options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED;
  return readXlsbPackage(inflateSpreadsheetPackage(data, cap));
}

/**
 * Build the model from an already-inflated `.xlsb` package. Separate from {@link readXlsb} so the
 * `.xlsx` reader can hand over a package it has already inflated and classified, rather than
 * inflating the same bytes twice.
 */
export function readXlsbPackage(files: Record<string, Uint8Array>): Workbook {
  const {partText, partBytes} = packageAccessors(files);
  const workbookPart = partBytes(XLSB_WORKBOOK_PART);
  if (workbookPart === undefined) {
    throw new UnsupportedFormatError(
      'unknown',
      `not a valid .xlsb package: ${XLSB_WORKBOOK_PART} is missing`,
    );
  }

  const rels = parseRelationships(partText('xl/_rels/workbook.bin.rels') ?? '');
  const sharedStrings = parseSharedStrings(partBytes('xl/sharedStrings.bin'));
  const {cellXfs, namedStyles} = parseStyleTable(partBytes('xl/styles.bin'));

  const workbook = new Workbook();
  // As in the XML reader, the named-style layer is restored only when a file declares more than the
  // Normal default, so an ordinary workbook keeps an empty table and writes just that default back.
  if (namedStyles.length > 1) workbook.restoreNamedStyles(namedStyles);

  for (const declared of readSheetDeclarations(workbookPart)) {
    const sheet = workbook.addWorksheet(declared.name, {state: declared.state});
    const target = declared.relId === undefined ? undefined : rels.get(declared.relId);
    const part = target === undefined ? undefined : partBytes(resolveWorkbookPart(target));
    if (part !== undefined) parseWorksheet(part, sheet, sharedStrings, cellXfs);
  }
  return workbook;
}

// One sheet as `xl/workbook.bin` declares it, in workbook (tab) order.
interface SheetDeclaration {
  readonly name: string;
  readonly relId: string | undefined;
  readonly state: WorksheetState['state'];
}

// `BrtBundleSh` ([MS-XLSB] 2.4.303) records, one per sheet, between a Begin/End pair — the binary
// spelling of `<sheets><sheet name state r:id/></sheets>`.
function readSheetDeclarations(part: Uint8Array): SheetDeclaration[] {
  const sheets: SheetDeclaration[] = [];
  let inBundle = false;
  for (const record of readRecords(part)) {
    if (record.type === BRT.BeginBundleShs) inBundle = true;
    else if (record.type === BRT.EndBundleShs) break;
    else if (record.type === BRT.BundleSh && inBundle) {
      const reader = new RecordReader(record.data);
      const state = SHEET_STATES[reader.u32()] ?? 'visible';
      reader.skip(4); // iTabId: the sheet's stable id, which the model assigns itself.
      const relId = reader.nullableWideString();
      sheets.push({name: reader.wideString(), relId, state});
    }
  }
  return sheets;
}

// `hsState` ([MS-XLSB] 2.4.303), indexed by its stored value.
const SHEET_STATES: ReadonlyArray<WorksheetState['state']> = ['visible', 'hidden', 'veryHidden'];
