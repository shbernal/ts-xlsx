// The buffered `.xlsb` reader: a binary OPC package in, a {@link Workbook} model out.
//
// An `.xlsb` is the *same* OPC/ZIP container and the *same* relationship graph as an `.xlsx` — the
// workbook part points at its sheets, its shared strings, and its style sheet through ordinary
// `.rels` XML. Only the office-document parts differ: `xl/workbook.bin`, `xl/worksheets/sheetN.bin`
// and friends are BIFF12 record streams instead of XML. So this module reuses the container layer
// wholesale (`../opc/`, the bounded inflater) and swaps only the part parsers, which is
// exactly the "two codecs over one model" shape the format note argues for.
//
// The model produced is the one `readXlsx` produces, not a parallel one: the same `Workbook`, the
// same `XfStyle` table, the same cells. That is what lets a caller convert between the two forms —
// and what the corpus asserts, by reading a workbook Excel saved in both forms and comparing.
//
// Not yet decoded (each its own slice of work, none silently wrong): rich-text runs, tables, pivots,
// and conditional formatting.

import {unmangleFunctions} from '../../core/formula.ts';
import {INTERNAL} from '../../core/internal.ts';
import {type DefinedName, Workbook} from '../../core/workbook.ts';
import type {WorksheetState} from '../../core/worksheet.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {packageAccessors, parseRelationships, resolveWorkbookPart} from '../opc/read-opc.ts';
import {DEFAULT_MAX_UNCOMPRESSED, type ReadXlsxOptions} from '../opc/read-options.ts';
import {inflateSpreadsheetPackage} from '../opc/sniff-format.ts';
import {decodeFormula, type ExternSheetRef, type FormulaScope} from './formula.ts';
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
  if (namedStyles.length > 1) workbook[INTERNAL].restoreNamedStyles(namedStyles);

  const declaration = readWorkbookPart(workbookPart);
  const scope: FormulaScope = {
    sheetNames: declaration.sheets.map((sheet) => sheet.name),
    externSheets: declaration.externSheets,
    selfSupBook: declaration.selfSupBook,
    names: declaration.names.map((name) => name.name),
  };

  for (const declared of declaration.sheets) {
    const sheet = workbook.addWorksheet(declared.name, {state: declared.state});
    const target = declared.relId === undefined ? undefined : rels.get(declared.relId);
    const part = target === undefined ? undefined : partBytes(resolveWorkbookPart(target));
    if (part !== undefined) parseWorksheet(part, sheet, sharedStrings, cellXfs, scope);
  }
  for (const defined of definedNames(declaration, scope)) workbook.defineName(defined);
  return workbook;
}

// One sheet as `xl/workbook.bin` declares it, in workbook (tab) order.
interface SheetDeclaration {
  readonly name: string;
  readonly relId: string | undefined;
  readonly state: WorksheetState['state'];
}

// One `BrtName`, still in its on-disk form: the target is a token stream that cannot be decoded until
// every sheet is known, so the record is gathered first and resolved after the pass.
interface NameDeclaration {
  readonly name: string;
  /** Zero-based index of the sheet the name is scoped to, or `undefined` for a workbook-global name. */
  readonly scopeSheet: number | undefined;
  /** Whether the name registers a callable rather than a target — see {@link definedNames}. */
  readonly isFunction: boolean;
  readonly rgce: Uint8Array;
  readonly rgcb: Uint8Array;
}

interface WorkbookDeclaration {
  readonly sheets: readonly SheetDeclaration[];
  readonly names: readonly NameDeclaration[];
  readonly externSheets: readonly ExternSheetRef[];
  readonly selfSupBook: number | undefined;
}

// One pass over `xl/workbook.bin`, gathering everything the rest of the read depends on: the sheet
// bundle, the externals block a 3-D reference resolves through, and the defined names.
function readWorkbookPart(part: Uint8Array): WorkbookDeclaration {
  const sheets: SheetDeclaration[] = [];
  const names: NameDeclaration[] = [];
  let externSheets: readonly ExternSheetRef[] = [];
  let inBundle = false;
  let inExternals = false;
  // A workbook with no external links declares exactly one supporting book: itself. Rather than
  // enumerate every record type that could open another — and risk miscounting into a *wrong* sheet
  // name — anything else inside the externals block disqualifies the whole table.
  let supportingBooks = 0;
  let selfSupBook: number | undefined;

  for (const record of readRecords(part)) {
    if (record.type === BRT.BeginBundleShs) inBundle = true;
    else if (record.type === BRT.EndBundleShs) inBundle = false;
    else if (record.type === BRT.BundleSh && inBundle) sheets.push(readSheet(record.data));
    else if (record.type === BRT.BeginExternals) inExternals = true;
    else if (record.type === BRT.EndExternals) inExternals = false;
    else if (record.type === BRT.ExternSheet) externSheets = readExternSheets(record.data);
    else if (record.type === BRT.SupSelf) selfSupBook = supportingBooks++;
    else if (inExternals) supportingBooks++;
    else if (record.type === BRT.Name) names.push(readName(record.data));
  }
  return {
    sheets,
    names,
    externSheets,
    selfSupBook: supportingBooks === 1 ? selfSupBook : undefined,
  };
}

// `BrtBundleSh` ([MS-XLSB] 2.4.303) — the binary spelling of `<sheet name state r:id/>`.
function readSheet(data: Uint8Array): SheetDeclaration {
  const reader = new RecordReader(data);
  const state = SHEET_STATES[reader.u32()] ?? 'visible';
  reader.skip(4); // iTabId: the sheet's stable id, which the model assigns itself.
  const relId = reader.nullableWideString();
  return {name: reader.wideString(), relId, state};
}

// `BrtExternSheet` ([MS-XLSB] 2.4.677): the `Xti` table every 3-D reference indexes into.
function readExternSheets(data: Uint8Array): ExternSheetRef[] {
  const reader = new RecordReader(data);
  const count = reader.u32();
  const entries: ExternSheetRef[] = [];
  // Each entry is three 4-byte fields; checking the count against what the record holds keeps a forged
  // one from driving the loop rather than the record's own length.
  if (count * XTI_BYTES > reader.remaining) return entries;
  for (let index = 0; index < count; index++) {
    entries.push({supBook: reader.u32(), firstSheet: reader.i32(), lastSheet: reader.i32()});
  }
  return entries;
}

const XTI_BYTES = 12;

// `BrtName` ([MS-XLSB] 2.4.673). The target is left undecoded here: it is a token stream that may cite
// a sheet by index, and the sheet bundle is not necessarily complete at this point in the stream.
function readName(data: Uint8Array): NameDeclaration {
  const reader = new RecordReader(data);
  const flags = reader.u32();
  reader.skip(1); // chKey: the Alt-key shortcut a macro name can carry.
  const itab = reader.u32();
  const name = reader.wideString();
  const rgce = reader.bytes(reader.u32());
  const rgcb = reader.bytes(reader.u32());
  return {
    name,
    scopeSheet: itab === GLOBAL_NAME_SCOPE ? undefined : itab,
    isFunction: (flags & NAME_IS_FUNCTION) !== 0,
    rgce,
    rgcb,
  };
}

// `itab` for a workbook-global name; any other value is a zero-based sheet index.
const GLOBAL_NAME_SCOPE = 0xffffffff;
// `fFunc`: the name registers something callable rather than a range.
const NAME_IS_FUNCTION = 0x00000002;

// The workbook's defined names, as the model holds them.
//
// Two kinds of `BrtName` are dropped, both because the XML form does not persist them either — so
// carrying them through would make the two readings of one workbook disagree. A *function* name is
// Excel's registration of a callable (every post-2007 function gets one, `_xlfn.TEXTJOIN` and
// friends); its target is the placeholder `#NAME?`, not a range. And `_xlnm._FilterDatabase` is the
// built-in Excel derives from a sheet's autofilter, which the model reconstructs from the autofilter
// itself. A name whose target uses a token this reader cannot decode is dropped too, rather than
// surfaced with a target that is a guess.
function definedNames(declaration: WorkbookDeclaration, scope: FormulaScope): DefinedName[] {
  const names: DefinedName[] = [];
  for (const declared of declaration.names) {
    if (declared.isFunction || declared.name === FILTER_DATABASE_NAME) continue;
    const refersTo = decodeFormula(declared.rgce, declared.rgcb, scope);
    if (refersTo === undefined) continue;
    const sheet =
      declared.scopeSheet === undefined ? undefined : scope.sheetNames[declared.scopeSheet];
    names.push({
      name: declared.name,
      ...(sheet === undefined ? {} : {scope: sheet}),
      // Stripped back to the readable form, the same normalisation the XML reader applies.
      refersTo: unmangleFunctions(refersTo),
    });
  }
  return names;
}

const FILTER_DATABASE_NAME = '_xlnm._FilterDatabase';

// `hsState` ([MS-XLSB] 2.4.303), indexed by its stored value.
const SHEET_STATES: ReadonlyArray<WorksheetState['state']> = ['visible', 'hidden', 'veryHidden'];
