# CSV

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CsvReadOptions`

<sub>interface</sub>

```ts
interface CsvReadOptions {
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
```

---

### `CsvWriteOptions`

<sub>interface</sub>

```ts
interface CsvWriteOptions {
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
```

---

### `readCsv`

<sub>function</sub>

Parse CSV text (or UTF-8 bytes) into a workbook holding a single worksheet.

```ts
function readCsv(input: string | Uint8Array, options: CsvReadOptions = {}): Workbook;
```

---

### `writeCsv`

<sub>function</sub>

The CSV bytes of one worksheet in the requested encoding, with a UTF-8 BOM by default.

```ts
function writeCsv(workbook: Workbook, options: CsvWriteOptions = {}): Uint8Array;
```

---

### `writeCsvText`

<sub>function</sub>

The logical CSV text of one worksheet — no BOM, no byte encoding.

```ts
function writeCsvText(workbook: Workbook, options: CsvWriteOptions = {}): string;
```
