# Streaming writes

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CalcProperties`

<sub>interface</sub>

Calculation settings applied to the streamed workbook. Mirrors the [`Workbook`](./workbook.md#workbook) flags.

```ts
interface CalcProperties {
  /** Ask the consumer to recalculate every formula on open — the OOXML `fullCalcOnLoad` flag. */
  fullCalcOnLoad?: boolean;
}
```

---

### `StreamedRow`

<sub>class</sub>

A row appended to a [`WorksheetStreamWriter`](./streaming-writes.md#worksheetstreamwriter). Style its cells through [`cells`](./streaming-writes.md#streamedrowcells), then call
[`commit`](./streaming-writes.md#streamedrowcommit) to mark it finished. In an eager (inline-strings) writer, committing serialises the
row and frees its cells from the model, bounding peak memory; with `useSharedStrings` on it is a
no-op and the row stays live until the workbook commits.

```ts
class StreamedRow {
  get cells(): readonly Cell[];
  commit(): void;
}
```

**Members**

#### `StreamedRow.cells`

```ts
get cells(): readonly Cell[];
```

The cells this row materialised, for styling before it is committed.

#### `StreamedRow.commit`

```ts
commit(): void;
```

Finalise the row: an eager writer serialises it now and releases its cells; otherwise a no-op.
Committing twice is harmless — the second call does nothing rather than re-emitting the row.

---

### `WorkbookStreamWriter`

<sub>class</sub>

A workbook written incrementally to a Node stream. Add worksheets, append their rows, commit each
sheet, then [`commit`](./streaming-writes.md#workbookstreamwritercommit) the workbook to assemble and stream the package. The produced bytes are
available both as the resolved value of `commit()` and through [`stream`](./streaming-writes.md#workbookstreamwriterstream) (a Node `Readable`
that a caller can `pipe`).

```ts
class WorkbookStreamWriter {
  readonly calcProperties: CalcProperties = {};
  get properties(): Workbook['properties'];
  get stream(): Readable;
  addImage(options: AddImageOptions): number;
  addWorksheet(name: string, options: AddWorksheetOptions = {}): WorksheetStreamWriter;
  async commit(): Promise<Uint8Array>;
}
```

**Members**

#### `WorkbookStreamWriter.calcProperties`

```ts
readonly calcProperties: CalcProperties = {};
```

Calculation settings for the workbook; set `fullCalcOnLoad` before committing to emit it.

#### `WorkbookStreamWriter.properties`

```ts
get properties(): Workbook['properties'];
```

Document-level metadata written to the package's core properties.

#### `WorkbookStreamWriter.stream`

```ts
get stream(): Readable;
```

The output stream carrying the package bytes. A caller drives it with Node's standard idiom —
`writer.stream.pipe(out)` — which composes because `pipe` returns its destination. The stream is
created lazily on first access so a caller handing the writer its own sink is still free to
ignore this one.

#### `WorkbookStreamWriter.addImage`

```ts
addImage(options: AddImageOptions): number;
```

Register a picture's bytes on the workbook's shared media registry and return its id, to anchor
on any sheet with [`WorksheetStreamWriter.addImage`](./streaming-writes.md#worksheetstreamwriteraddimage). Mirrors [`Workbook.addImage`](./workbook.md#workbookaddimage): one
media part backs an image anchored on several sheets. Rejected once the workbook is committed.

#### `WorkbookStreamWriter.addWorksheet`

```ts
addWorksheet(name: string, options: AddWorksheetOptions = {}): WorksheetStreamWriter;
```

Create a worksheet and append it to the workbook.

#### `WorkbookStreamWriter.commit`

```ts
async commit(): Promise<Uint8Array>;
```

Assemble the workbook into its package, stream the bytes through [`stream`](./streaming-writes.md#workbookstreamwriterstream), and resolve with
the same bytes. Every sheet is frozen first, so a row added after this rejects legibly. Idempotent
only in that a second call throws rather than re-emitting.

---

### `WorkbookStreamWriterOptions`

<sub>type</sub>

Options fixed at construction that shape the whole streamed package.

```ts
type WorkbookStreamWriterOptions = SinkOptions & {
  /**
   * Pool plain string cell values into a shared-strings table rather than storing each inline — the
   * same {@link WriteOptions.useSharedStrings} the buffered writer exposes. Off by default.
   */
  readonly useSharedStrings?: boolean;
};
```

---

### `WorksheetStreamWriter`

<sub>class</sub>

A worksheet being written incrementally. Append rows with [`addRow`](./streaming-writes.md#worksheetstreamwriteraddrow)/[`addRows`](./streaming-writes.md#worksheetstreamwriteraddrows), style
cells through [`getCell`](./streaming-writes.md#worksheetstreamwritergetcell), then [`commit`](./streaming-writes.md#worksheetstreamwritercommit) to freeze it — after which any further mutation
is rejected with a legible error rather than silently accepted or crashing.

```ts
class WorksheetStreamWriter {
  get name(): string;
  get rowCount(): number;
  addRow(values: CellValue[]): StreamedRow;
  addRows(rows: CellValue[][]): StreamedRow[];
  flushRow(number: number, cells: readonly Cell[]): void;
  flushedSheet(): FlushedSheet | undefined;
  getCell(reference: string): Cell;
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
  addConditionalFormatting(formatting: ConditionalFormatting): void;
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void;
  set autoFilter(filter: string | AutoFilter | undefined);
  get autoFilter(): AutoFilter | undefined;
  protect(password?: string, options: SheetProtectionOptions = {}): void;
  commit(): void;
  get committed(): boolean;
  get model(): Worksheet;
}
```

**Members**

#### `WorksheetStreamWriter.name`

```ts
get name(): string;
```

The sheet's name.

#### `WorksheetStreamWriter.rowCount`

```ts
get rowCount(): number;
```

The number of rows written so far — spans gaps and formatted-only rows, like the model, and
survives the eviction of eagerly-flushed rows.

#### `WorksheetStreamWriter.addRow`

```ts
addRow(values: CellValue[]): StreamedRow;
```

Append one row of values after the last used row; the cells are returned for styling.

#### `WorksheetStreamWriter.addRows`

```ts
addRows(rows: CellValue[][]): StreamedRow[];
```

Append a batch of rows in one call, each landing directly below the previous.

#### `WorksheetStreamWriter.flushRow`

```ts
flushRow(number: number, cells: readonly Cell[]): void;
```

Serialise an eagerly-committed row and release its cells from the model. Called by
[`StreamedRow.commit`](./streaming-writes.md#streamedrowcommit); the row's `<row>` XML is retained (interned into the workbook's live
style registry so its ids stay valid) and the cell graph is dropped, bounding peak memory.

**Throws** — [`AuthoringError`](./errors.md#authoringerror) if the row carries a shared-formula cell — a finished row cannot join the
whole-sheet formula planning, so shared formulas must be authored through [`getCell`](./streaming-writes.md#worksheetstreamwritergetcell).

#### `WorksheetStreamWriter.getCell`

```ts
getCell(reference: string): Cell;
```

Address a cell by its A1 reference to read or style it before the sheet is committed.

#### `WorksheetStreamWriter.addDataValidation`

```ts
addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
```

Attach a data validation to a range before the sheet is committed. Delegates to the model, so the
streamed package emits the `<dataValidations>` block in its CT_Worksheet position — before
`<hyperlinks>` — because both writers share one worksheet serializer.

#### `WorksheetStreamWriter.addConditionalFormatting`

```ts
addConditionalFormatting(formatting: ConditionalFormatting): void;
```

Attach a conditional formatting to a range before the sheet is committed. Like every other block,
it lands in its schema-mandated slot — after `<mergeCells>`, before `<dataValidations>` and
`<hyperlinks>` — since the streamed sheet is serialized through the same path as a buffered write.

#### `WorksheetStreamWriter.addImage`

```ts
addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void;
```

Anchor a workbook image (the id from [`WorkbookStreamWriter.addImage`](./streaming-writes.md#workbookstreamwriteraddimage)) to this sheet,
spanning the rectangle from the top-left grid point `tl` to the bottom-right `br`. The streamed
package emits the drawing part, its media relationship, and the sheet's `<drawing>` reference
exactly as a buffered write does — both writers share `buildPackageParts`.

#### `WorksheetStreamWriter.autoFilter`

```ts
set autoFilter(filter: string | AutoFilter | undefined);
get autoFilter(): AutoFilter | undefined;
```

Apply the sheet's autofilter before it is committed; mirrors [`Worksheet.autoFilter`](./worksheet.md#worksheetautofilter). The
streamed package emits `<autoFilter>` in its CT_Worksheet slot — after `<sheetProtection>` — and
contributes the hidden `_FilterDatabase` defined name, exactly as a buffered write does.

#### `WorksheetStreamWriter.protect`

```ts
protect(password?: string, options: SheetProtectionOptions = {}): void;
```

Apply sheet-level protection before the sheet is committed; mirrors [`Worksheet.protect`](./worksheet.md#worksheetprotect). The
shared serializer places `<sheetProtection>` ahead of `<autoFilter>` per CT_Worksheet, so a
streamed sheet carrying both stays valid rather than corrupt.

#### `WorksheetStreamWriter.commit`

```ts
commit(): void;
```

Freeze the sheet: no more rows or edits may be added after this.

#### `WorksheetStreamWriter.committed`

```ts
get committed(): boolean;
```

Whether the sheet has been committed.
