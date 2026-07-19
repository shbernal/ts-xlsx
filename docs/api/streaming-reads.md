# Streaming reads

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `readSheetRows`

<sub>function</sub>

Stream a worksheet's rows from an `.xlsx` package, yielding each in sheet order without building
the workbook model. Only rows the sheet actually declares are yielded, and within a row only its
non-empty cells — a blank or style-only cell contributes nothing, matching the intent of a data
read.

```ts
function* readSheetRows(data: Uint8Array, options: ReadSheetRowsOptions = {}): Generator<StreamedRow>;
```

- `data` — The raw `.xlsx` bytes.
- `options` — Sheet selector and the inflate bound (see `ReadSheetRowsOptions`).
**Throws** — if the archive is malformed, exceeds the inflate bound, or names no worksheet —
or if `options.sheet` selects a sheet that does not exist.

---

### `ReadSheetRowsOptions`

<sub>interface</sub>

```ts
interface ReadSheetRowsOptions extends ReadXlsxOptions {
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
function* readWorkbookStream(data: Uint8Array, options: ReadXlsxOptions = {}): Generator<StreamedSheet>;
```

- `data` — The raw `.xlsx` bytes.
- `options` — The inflate bound (see `ReadXlsxOptions`).
**Throws** — if the archive is malformed or exceeds the inflate bound.

---

### `StreamedCell`

<sub>interface</sub>

One non-empty cell in a `StreamedRow`.

```ts
interface StreamedCell {
    readonly col: number;
    readonly address: string;
    readonly value: CellValue;
    readonly style?: StreamedCellStyle;
}
```

---

### `StreamedCellStyle`

<sub>type</sub>

The resolved style facets of a streamed cell — its own `<c s>` cell format, flattened exactly as
the buffered reader resolves it. Present only when the cell carries a format; a consumer can copy
these straight onto a writer cell to preserve its look through a streaming read→write.

```ts
type StreamedCellStyle = XfStyle;
```

---

### `StreamedRow`

<sub>interface</sub>

One worksheet row, as yielded by `readSheetRows` / `StreamedSheet.rows`.

```ts
interface StreamedRow {
    readonly number: number;
    readonly hidden: boolean;
    readonly cells: readonly StreamedCell[];
}
```

---

### `StreamedSheet`

<sub>interface</sub>

One worksheet, as yielded by `readWorkbookStream`. The sheet's `rows` stream one at a
time; its `hiddenColumns` and `merges` are populated by that same single pass.

The two summaries are resolved lazily: reading either accessor drives a full scan of the sheet if
its rows have not already been consumed, so their order relative to `rows()` never matters. (When
rows *are* consumed first — the streaming idiom — the accessors reuse that pass and re-scan
nothing.)

```ts
interface StreamedSheet {
    readonly name: string;
    rows(): Generator<StreamedRow>;
    readonly hiddenColumns: readonly number[];
    readonly merges: readonly string[];
}
```
