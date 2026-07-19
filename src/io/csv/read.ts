// CSV parsing — flat delimited text back into a one-sheet workbook.
//
// The hard part of reading CSV is not splitting fields; it is deciding a field's *type* without
// corrupting data. The rules here are deliberate and lossless-by-default:
//   - An empty field is the empty cell (`null`); a whitespace-only field is a string, never the
//     number 0 that `Number("   ")` would silently produce.
//   - A numeric-looking field becomes a number only when its magnitude is exactly representable
//     (within the safe-integer range); an oversized id like a 20-digit account number is kept as
//     its original string so no digits are lost.
//   - Only a strictly-formatted ISO date (`YYYY-MM-DD`, optional time) becomes a Date; padded ids
//     and dash-codes such as `2020-00001` or `1-3` stay strings.
// A caller can override coercion wholesale with `map` (e.g. the identity function to keep every
// field a raw string, preserving leading zeros).

import type {CellValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';

export interface CsvReadOptions {
  /** Field separator; defaults to a comma. A single character. */
  readonly delimiter?: string;
  /** Treat the first line as a header and drop it, leaving only data rows. */
  readonly headers?: boolean;
  /** Per-field transform replacing the default type coercion; receives the raw string and its
   * 0-based column index. */
  readonly map?: (value: string, index: number) => CellValue;
  /** Name for the single worksheet produced; defaults to `"Sheet1"`. */
  readonly sheetName?: string;
}

/** Parse CSV text (or UTF-8 bytes) into a workbook holding a single worksheet. */
export function readCsv(input: string | Uint8Array, options: CsvReadOptions = {}): Workbook {
  const text = stripBom(typeof input === 'string' ? input : Buffer.from(input).toString('utf8'));
  const delimiter = options.delimiter ?? ',';

  let rows = parseCsvRows(text, delimiter);
  if (options.headers) rows = rows.slice(1);

  const coerce = options.map ?? defaultCsvCoerce;
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet(options.sheetName ?? 'Sheet1');
  sheet.addRows(rows.map((fields) => fields.map((field, index) => coerce(field, index))));
  return workbook;
}

function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

// A character-scan parser: quotes toggle literal mode, a doubled quote inside a quoted field is one
// quote, and a row ends on LF (a preceding CR is dropped, so CRLF and LF read the same). A final
// trailing newline does not yield a spurious empty row.
function parseCsvRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else if (ch !== '\r') {
      field += ch;
    }
  }

  row.push(field);
  const trailingNewline = text.endsWith('\n') || text.endsWith('\r');
  if (!(trailingNewline && row.length === 1 && row[0] === '')) rows.push(row);
  return rows;
}

const NUMERIC = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/;

function defaultCsvCoerce(field: string): CellValue {
  if (field === '') return null;
  if (NUMERIC.test(field)) {
    const value = Number(field);
    // Beyond the safe-integer range a double silently loses digits; keep the original text instead.
    return Number.isFinite(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER ? value : field;
  }
  const iso = ISO_DATE.exec(field);
  if (iso) {
    const date = isoToDate(iso);
    if (date !== null) return date;
  }
  return field;
}

function isoToDate(match: RegExpExecArray): Date | null {
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = match[4] === undefined ? 0 : Number(match[4]);
  const minute = match[5] === undefined ? 0 : Number(match[5]);
  const second = match[6] === undefined ? 0 : Number(match[6]);
  if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59 || second > 59) {
    return null;
  }
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  return Number.isNaN(date.getTime()) ? null : date;
}
