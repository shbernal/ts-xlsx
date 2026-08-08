// CSV serialization — the flat-text sibling of the XLSX writer.
//
// A worksheet is a rectangle of typed cells; CSV is that rectangle flattened to delimited
// text. The lossy direction (styles, formulas-as-formulas, multiple sheets) is inherent to the
// format, so this writer makes the honest choices explicit: one selected sheet, each row sized to
// its own populated extent (never clamped to a sibling row's width), a formula rendered as its
// cached result, a Date rendered by a caller-supplied format or a full ISO-8601 timestamp. What a
// value reads as is `cellValueToText`'s answer, not a private one — a CSV field and `cell.text`
// disagreeing about the same cell would be a bug in one of them.
//
// `writeCsvText` yields the logical text; `writeCsv` encodes it to bytes and — for UTF-8, the
// default — prepends a byte-order mark so a consumer such as Excel detects the encoding and does
// not mangle non-ASCII on open. The BOM is a byte-level marker, not part of the logical text.

import type {Cell} from '../../core/cell.ts';
import type {CellValue} from '../../core/value.ts';
import {
  cellValueToText,
  isDataTableFormulaValue,
  isFormulaValue,
  isSharedFormulaValue,
} from '../../core/value.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';

export interface CsvWriteOptions {
  /** Which worksheet to write; defaults to the first. A name matching no sheet throws rather than
   * silently emitting an empty file. */
  readonly sheetName?: string;
  /** Field separator; defaults to a comma. */
  readonly delimiter?: string;
  /** Line separator between rows; defaults to `"\n"`. */
  readonly rowDelimiter?: string;
  /** A token format (e.g. `"MM/DD/YYYY"`) for Date cells; without it a Date renders as a full
   * ISO-8601 timestamp. */
  readonly dateFormat?: string;
  /** Render Date cells in UTC rather than the runner's local time. */
  readonly dateUTC?: boolean;
  /** Byte encoding for {@link writeCsv}; defaults to `"utf8"`. */
  readonly encoding?: BufferEncoding;
  /** Prepend a UTF-8 byte-order mark (applies only to UTF-8); defaults to `true` for UTF-8. */
  readonly bom?: boolean;
  /** Per-field transform replacing the default value rendering; receives the cell's value (`null`
   * for an unpopulated column) and its 0-based column index. Quoting (commas, quotes, newlines) is
   * still applied to the returned text. */
  readonly map?: (value: CellValue, index: number) => string;
}

const UTF8_BOM = Uint8Array.of(0xef, 0xbb, 0xbf);

/** The logical CSV text of one worksheet — no BOM, no byte encoding. */
export function writeCsvText(workbook: Workbook, options: CsvWriteOptions = {}): string {
  const sheet = selectSheet(workbook, options.sheetName);
  const delimiter = options.delimiter ?? ',';
  const rowDelimiter = options.rowDelimiter ?? '\n';

  const lines: string[] = [];
  for (const {cells} of sheet.rows()) {
    let width = 0;
    const byColumn = new Map<number, Cell>();
    for (const cell of cells) {
      byColumn.set(cell.col, cell);
      if (cell.col > width) width = cell.col;
    }
    const fields: string[] = [];
    for (let column = 1; column <= width; column++) {
      const value = byColumn.get(column)?.value ?? null;
      const text = options.map ? options.map(value, column - 1) : csvFieldText(value, options);
      fields.push(quoteField(text, delimiter));
    }
    lines.push(fields.join(delimiter));
  }
  return lines.join(rowDelimiter);
}

/** The CSV bytes of one worksheet in the requested encoding, with a UTF-8 BOM by default. */
export function writeCsv(workbook: Workbook, options: CsvWriteOptions = {}): Uint8Array {
  const text = writeCsvText(workbook, options);
  const encoding = options.encoding ?? 'utf8';
  const body = Buffer.from(text, encoding);
  const wantBom = options.bom ?? encoding === 'utf8';
  if (!wantBom || encoding !== 'utf8') return Uint8Array.from(body);

  const out = new Uint8Array(UTF8_BOM.length + body.length);
  out.set(UTF8_BOM, 0);
  out.set(body, UTF8_BOM.length);
  return out;
}

function selectSheet(workbook: Workbook, name: string | undefined): Worksheet {
  if (name === undefined) {
    const first = workbook.worksheets[0];
    if (first === undefined) throw new AuthoringError('workbook has no worksheet to write as CSV');
    return first;
  }
  const sheet = workbook.getWorksheet(name);
  if (sheet === undefined) throw new AuthoringError(`no worksheet named "${name}" to write as CSV`);
  return sheet;
}

// A field is the value's plain text, with one CSV-only deviation: `dateFormat`/`dateUTC` let a
// caller render dates in something other than ISO-8601. That reaches inside a formula's cached
// result too, which is why the recursion is here rather than delegated wholesale.
function csvFieldText(value: CellValue, options: CsvWriteOptions): string {
  if (value instanceof Date && options.dateFormat !== undefined) {
    return formatDate(value, options.dateFormat, options.dateUTC ?? false);
  }
  if (isFormulaValue(value) || isSharedFormulaValue(value) || isDataTableFormulaValue(value)) {
    return value.result === undefined ? '' : csvFieldText(value.result, options);
  }
  return cellValueToText(value);
}

const DATE_TOKENS = /YYYY|YY|MM|DD|HH|mm|ss|M|D|H|m|s/g;

function formatDate(date: Date, format: string | undefined, utc: boolean): string {
  if (Number.isNaN(date.getTime())) return '';
  if (format === undefined) return date.toISOString();

  const year = utc ? date.getUTCFullYear() : date.getFullYear();
  const month = (utc ? date.getUTCMonth() : date.getMonth()) + 1;
  const day = utc ? date.getUTCDate() : date.getDate();
  const hour = utc ? date.getUTCHours() : date.getHours();
  const minute = utc ? date.getUTCMinutes() : date.getMinutes();
  const second = utc ? date.getUTCSeconds() : date.getSeconds();
  const pad = (n: number): string => String(n).padStart(2, '0');
  const tokens: Record<string, string> = {
    YYYY: String(year),
    YY: pad(year % 100),
    MM: pad(month),
    M: String(month),
    DD: pad(day),
    D: String(day),
    HH: pad(hour),
    H: String(hour),
    mm: pad(minute),
    m: String(minute),
    ss: pad(second),
    s: String(second),
  };
  return format.replace(DATE_TOKENS, (token) => tokens[token] ?? token);
}

function quoteField(field: string, delimiter: string): string {
  if (
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes('\n') ||
    field.includes('\r')
  ) {
    return `"${field.replace(/"/g, '""')}"`;
  }
  return field;
}
