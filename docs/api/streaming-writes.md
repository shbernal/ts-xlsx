# Streaming writes

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AddStreamedWorksheetOptions`

<sub>interface</sub>

Options for [`WorkbookStreamWriter.addWorksheet`](./streaming-writes.md#workbookstreamwriteraddworksheet).

```ts
interface AddStreamedWorksheetOptions extends AddWorksheetOptions {
  /**
   * The deepest row outline level the sheet will use, written as `<sheetFormatPr outlineLevelRow>`.
   * That element precedes the rows, so a streamed sheet writes it when its first row is committed,
   * from the larger of this and the deepest level among the rows written then. A row committed later
   * with a deeper level is refused, so declare the depth here when the outline starts further down
   * the sheet. A consumer sizes the outline bar from it.
   */
  readonly outlineLevelRow?: number;
}
```

---

### `CalcProperties`

<sub>interface</sub>

Calculation settings applied to the streamed workbook. Mirrors the [`Workbook`](./workbook.md#workbook) flags.

```ts
interface CalcProperties {
  /** Ask the consumer to recalculate every formula on open: the OOXML `fullCalcOnLoad` flag. */
  fullCalcOnLoad?: boolean;
}
```

---

### `StreamedRow`

<sub>class</sub>

A row appended to a [`WorksheetStreamWriter`](./streaming-writes.md#worksheetstreamwriter). Style its cells through [`cells`](./streaming-writes.md#streamedrowcells), then call
[`commit`](./streaming-writes.md#streamedrowcommit) to mark it finished. In an eager (inline-strings) writer, committing writes the row
into the sheet's zip entry and frees its cells; with `useSharedStrings` on it is a no-op and the row
stays live until the workbook commits.

```ts
class StreamedRow {
  get cells(): readonly Cell[];
  commit(): Promise<void>;
}
```

**Members**

#### `StreamedRow.cells`

```ts
get cells(): readonly Cell[];
```

The cells `addRow` materialised, for styling before it is committed. A cell added to the row
through [`WorksheetStreamWriter.getCell`](./streaming-writes.md#worksheetstreamwritergetcell) is not among them, and is committed all the same.

#### `StreamedRow.commit`

```ts
commit(): Promise<void>;
```

Finalise the row. An eager writer writes it, and every uncommitted row above it, into the sheet's
zip entry and releases their cells; otherwise this is a no-op. Committing twice, or committing a
row a later row's commit already wrote, does nothing.

The promise resolves once the caller's sink can take more, so a producer that awaits it is paced
by the sink. It never rejects: a sink that fails is reported by [`WorkbookStreamWriter.commit`](./streaming-writes.md#workbookstreamwritercommit).

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if a row being written carries a shared-formula cell, a row outline level
deeper than the sheet's head declares, or an array formula whose range collides with another
formula. See `WorksheetStreamWriterInternals.commitThrough`.

---

### `WorkbookStreamWriter`

<sub>class</sub>

A workbook written incrementally to a Node stream. Add worksheets, append and commit their rows,
commit each sheet, then [`commit`](./streaming-writes.md#workbookstreamwritercommit) the workbook to write the rest of the package. With strings
inline, the default, a sheet's rows reach the destination as they are committed (ADR 0046).

```ts
class WorkbookStreamWriter {
  readonly calcProperties: CalcProperties = {};
  get properties(): Workbook['properties'];
  get stream(): Readable;
  addImage(options: AddImageOptions): number;
  addWorksheet(name: string, options: AddStreamedWorksheetOptions = {}): WorksheetStreamWriter;
  async commit(): Promise<Uint8Array | undefined>;
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

The output stream carrying the package bytes. A caller drives it with Node's standard idiom,
`writer.stream.pipe(out)`, which composes because `pipe` returns its destination. Created lazily
on first access, and from then on [`commit`](./streaming-writes.md#workbookstreamwritercommit) also resolves with the whole package, since this
stream may never be drained. It is not paced: the writer only waits on a sink it was given.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the writer has a sink and has already written bytes to it: those bytes
were never kept, so the stream would begin part-way through the archive. Take it before the
first row is committed.

#### `WorkbookStreamWriter.addImage`

```ts
addImage(options: AddImageOptions): number;
```

Register a picture's bytes on the workbook's shared media registry and return its id, to anchor
on any sheet with [`WorksheetStreamWriter.addImage`](./streaming-writes.md#worksheetstreamwriteraddimage). Mirrors [`Workbook.addImage`](./workbook.md#workbookaddimage): one
media part backs an image anchored on several sheets. Rejected once the workbook is committed.

#### `WorkbookStreamWriter.addWorksheet`

```ts
addWorksheet(name: string, options: AddStreamedWorksheetOptions = {}): WorksheetStreamWriter;
```

Create a worksheet and append it to the workbook. Its part is written into the archive as its rows
commit, so to keep memory bounded commit each sheet before writing the next: the archive holds one
entry at a time, and a later sheet's bytes wait in memory until the earlier one closes.

#### `WorkbookStreamWriter.commit`

```ts
async commit(): Promise<Uint8Array | undefined>;
```

Commit every sheet still open, write the rest of the package, and end the destinations. A second
call throws rather than re-emitting.

**What it resolves with, and why it can be `undefined`.** The archive is handed back only when
somebody asked for it: when no sink was supplied, or when [`stream`](./streaming-writes.md#workbookstreamwriterstream) was touched, since a
`PassThrough` nobody drained would otherwise be the only copy. A caller who passed `{filename}`
or their own `stream` and never reached for [`stream`](./streaming-writes.md#workbookstreamwriterstream) gets `undefined`, and the package is
never held as a whole: that is the memory the caller passed a sink to avoid. The type says so
rather than the prose alone, because a promise that resolves with a value the caller was told to
ignore is a promise they will use.

If assembling the package fails, or the sink failed at any point since the first row was written,
the returned promise rejects *and* every stream this writer was given or handed out is destroyed
with that error. A caller piping [`stream`](./streaming-writes.md#workbookstreamwriterstream), or awaiting its own sink, therefore sees an
`error` rather than waiting on a stream that will never end; and it is an `error` rather than a
clean `end` because the bytes written so far are a truncated package that must not be read as a
whole one.

---

### `WorkbookStreamWriterOptions`

<sub>type</sub>

Options fixed at construction that shape the whole streamed package.

```ts
type WorkbookStreamWriterOptions = SinkOptions & {
  /**
   * Pool plain string cell values into a shared-strings table rather than storing each inline: the
   * same {@link WriteOptions.useSharedStrings} the buffered writer exposes. Off by default.
   */
  readonly useSharedStrings?: boolean;
  /**
   * Which date system this workbook's serials count in ({@link Workbook.dateEpoch}). A construction
   * option rather than a settable property, unlike on the buffered writer: an eager writer serialises
   * each row as it is committed, so a system changed part-way through would leave the rows before the
   * change counting from a different day than the rows after it.
   */
  readonly dateEpoch?: DateEpoch;
};
```

---

### `WorksheetStreamWriter`

<sub>class</sub>

A worksheet being written incrementally. Append rows with [`addRow`](./streaming-writes.md#worksheetstreamwriteraddrow)/[`addRows`](./streaming-writes.md#worksheetstreamwriteraddrows), style
cells through [`getCell`](./streaming-writes.md#worksheetstreamwritergetcell), then [`commit`](./streaming-writes.md#worksheetstreamwritercommit) to finish it, after which any further mutation
is rejected with a legible error rather than silently accepted or crashing.

```ts
class WorksheetStreamWriter {
  get name(): string;
  get rowCount(): number;
  addRow(values: CellValue[]): StreamedRow;
  addRows(rows: CellValue[][]): StreamedRow[];
  getCell(reference: string): Cell;
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
  addHyperlink(link: Hyperlink): void;
  addConditionalFormatting(formatting: ConditionalFormatting): void;
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void;
  set autoFilter(filter: string | AutoFilter | undefined);
  get autoFilter(): AutoFilter | undefined;
  protect(password?: string, options: SheetProtectionOptions = {}): void;
  commit(): Promise<void>;
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

The number of rows written so far. Spans gaps and formatted-only rows, like the model, and
survives the eviction of written rows.

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

#### `WorksheetStreamWriter.getCell`

```ts
getCell(reference: string): Cell;
```

Address a cell by its A1 reference to read or style it before its row is written. Rows are written
in order, so the row must be below the last one written.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the reference names a row at or above the last row written. That row,
and every row above it, is already in the file and its cells are released, so the cell this
would materialise could only be written as a second row out of order.

#### `WorksheetStreamWriter.addDataValidation`

```ts
addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
```

Attach a data validation to a range before the sheet is committed. Delegates to the model, so the
streamed package emits the `<dataValidations>` block in its CT_Worksheet position, before
`<hyperlinks>`, because both writers share one worksheet serializer.

#### `WorksheetStreamWriter.addHyperlink`

```ts
addHyperlink(link: Hyperlink): void;
```

Put a hyperlink on a cell or a rectangle of cells before the sheet is committed; mirrors
[`Worksheet.addHyperlink`](./worksheet.md#worksheetaddhyperlink). A link lives beside the grid rather than in a row, so it may cover
cells of a row that is already committed.

#### `WorksheetStreamWriter.addConditionalFormatting`

```ts
addConditionalFormatting(formatting: ConditionalFormatting): void;
```

Attach a conditional formatting to a range before the sheet is committed. Like every other block,
it lands in its schema-mandated slot, after `<mergeCells>` and before `<dataValidations>` and
`<hyperlinks>`, since the streamed sheet is serialized through the same path as a buffered write.

#### `WorksheetStreamWriter.addImage`

```ts
addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void;
```

Anchor a workbook image (the id from [`WorkbookStreamWriter.addImage`](./streaming-writes.md#workbookstreamwriteraddimage)) to this sheet,
spanning the rectangle from the top-left grid point `tl` to the bottom-right `br`. The streamed
package emits the drawing part, its media relationship, and the sheet's `<drawing>` reference
exactly as a buffered write does.

#### `WorksheetStreamWriter.autoFilter`

```ts
set autoFilter(filter: string | AutoFilter | undefined);
get autoFilter(): AutoFilter | undefined;
```

Apply the sheet's autofilter before it is committed; mirrors [`Worksheet.autoFilter`](./worksheet.md#worksheetautofilter). The
streamed package emits `<autoFilter>` in its CT_Worksheet slot, after `<sheetProtection>`, and
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
commit(): Promise<void>;
```

Finish the sheet: an eager writer writes the rows still uncommitted and the rest of the part, and
closes its zip entry. No more rows or edits may be added after this; committing again returns the
same promise. The promise resolves as [`StreamedRow.commit`](./streaming-writes.md#streamedrowcommit)'s does.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if a row still to be written is refused (see [`StreamedRow.commit`](./streaming-writes.md#streamedrowcommit)),
if a column, the view or a sheet property changed after the first row was written, or if the
sheet's merges overlap a table.

#### `WorksheetStreamWriter.committed`

```ts
get committed(): boolean;
```

Whether the sheet has been committed.

#### `WorksheetStreamWriter.model`

```ts
get model(): Worksheet;
```

The underlying model sheet: row properties (an outline level, a height), column definitions and the
view are set here. What precedes the rows in the file (columns, view, sheet properties) is read
when the first row is written; set it before then.
