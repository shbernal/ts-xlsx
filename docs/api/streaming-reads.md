# Streaming reads

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readSheetRows`

<sub>function</sub>

Stream a worksheet's rows from an `.xlsx` package, yielding each in sheet order without building
the workbook model. Only rows the sheet actually declares are yielded, and within a row only its
non-empty cells — a blank or style-only cell contributes nothing, matching the intent of a data
read.

```ts
function* readSheetRows(
  data: Uint8Array,
  options: ReadSheetRowsOptions = {},
): Generator<StreamedRow>;
```

- `data` — The raw `.xlsx` bytes.
- `options` — Sheet selector and the inflate bound (see `ReadSheetRowsOptions`).
**Throws** — `UnsupportedFormatError` if the input is not a readable `.xlsx` package (a legacy `.xls`, a
binary `.xlsb`, or an unrecognised/non-ZIP blob — branch on `.format`).
**Throws** — `PackageReadError` if the input is a ZIP that cannot be unpacked — a corrupt or
truncated archive, or one exceeding the inflate bound (a probable zip bomb).
**Throws** — `XlsxParseError` if the package's workbook part declares no worksheets.
**Throws** — `RangeError` / `AuthoringError` if `options.sheet` selects a position, or a name,
that no worksheet has.

---

### `ReadSheetRowsOptions`

<sub>interface</sub>

```ts
interface ReadSheetRowsOptions extends ReadXlsxOptions {
  /**
   * Which worksheet to stream: its name, or its 1-based position in the workbook. Defaults to the
   * first sheet.
   */
  readonly sheet?: string | number;
}
```

---

### `readWorkbookStream`

<sub>function</sub>

Stream every worksheet of an `.xlsx` package in workbook order, without building the workbook
model. Each yielded `StreamedSheet` carries the declared sheet name and lets the caller
stream that sheet's rows and read its hidden-column and merge summaries — the streaming analogue
of walking `readXlsx(data).worksheets`.

```ts
function* readWorkbookStream(
  data: Uint8Array,
  options: ReadXlsxOptions = {},
): Generator<StreamedSheet>;
```

- `data` — The raw `.xlsx` bytes.
- `options` — The inflate bound (see `ReadXlsxOptions`).
**Throws** — `UnsupportedFormatError` if the input is not a readable `.xlsx` package (a legacy `.xls`, a
binary `.xlsb`, or an unrecognised/non-ZIP blob — branch on `.format`).
**Throws** — `PackageReadError` if the input is a ZIP that cannot be unpacked — a corrupt or
truncated archive, or one exceeding the inflate bound (a probable zip bomb).
