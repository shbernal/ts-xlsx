# CSV

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `CsvReadOptions`

<sub>interface</sub>

```ts
interface CsvReadOptions {
    readonly delimiter?: string;
    readonly headers?: boolean;
    readonly map?: (value: string, index: number) => CellValue;
    readonly sheetName?: string;
}
```

---

### `CsvWriteOptions`

<sub>interface</sub>

```ts
interface CsvWriteOptions {
    readonly sheetName?: string;
    readonly delimiter?: string;
    readonly rowDelimiter?: string;
    readonly dateFormat?: string;
    readonly dateUTC?: boolean;
    readonly encoding?: BufferEncoding;
    readonly bom?: boolean;
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
