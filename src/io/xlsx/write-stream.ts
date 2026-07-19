// The streaming `.xlsx` writer: author a workbook incrementally and emit its package through a Node
// stream, rather than holding the finished bytes in one buffer as {@link writeXlsx} does.
//
// A producer adds worksheets, appends rows as it generates them, commits each sheet, then commits the
// workbook — at which point the package is assembled and streamed out. The output goes to whichever
// destination the caller chose at construction: its own pull-based `stream` to pipe, a caller-owned
// `Writable` sink (an outbound upload), or a `filename` the writer opens. In every case `commit()`
// settles — it resolves once a supplied sink has flushed the whole package and rejects if that sink
// errors (an unopenable file), never hanging on a finish signal that will not come. The output rides a
// genuinely
// streamed zip container (fflate's `Zip`/`ZipDeflate`), which computes each entry's CRC-32
// incrementally, so the archive is well-formed by construction — the defect the upstream "streaming
// writer emits a corrupt zip" reports describe is structurally absent here. The bytes reload
// identically to a whole-file write because both writers share `buildPackageParts` for every part.
//
// Peak memory: a row committed with `row.commit()` is serialised to its `<row>` XML immediately and
// its cells evicted from the model, so an append-driven producer holds only the rows still in flight
// rather than the whole sheet's cell graph. The eagerly-rendered rows intern into the workbook's live
// style registry — the same one that emits `xl/styles.xml` — so their style ids stay correct; their
// XML is handed to `buildPackageParts` and emitted at the head of `<sheetData>`. This eager path runs
// with strings inline (a shared-strings pool is inherently whole-workbook, so it defeats bounding);
// turning `useSharedStrings` on falls back to holding every row live until commit. A flushed row is a
// finished row: it cannot join whole-sheet derivations, so a shared-formula clone in a committed row
// is rejected, and rows reached only through `getCell` (never `row.commit()`) stay live and serialise
// the ordinary way. The package bytes themselves are still assembled once at commit — a later slice
// can flush each sheet's `<sheetData>` straight into its streamed zip entry to bound that half too.

import {createWriteStream} from 'node:fs';
import {PassThrough, type Readable, type Writable} from 'node:stream';

import {Zip, ZipDeflate} from 'fflate';

import {encodeAddress} from '../../core/address.ts';
import type {AutoFilter} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import type {ConditionalFormatting} from '../../core/conditional-formatting.ts';
import type {DataValidation} from '../../core/data-validation.ts';
import type {AnchorPoint} from '../../core/image.ts';
import type {SheetProtectionOptions} from '../../core/protection.ts';
import {isSharedFormulaValue, type CellValue} from '../../core/value.ts';
import {Workbook, type AddImageOptions, type AddWorksheetOptions} from '../../core/workbook.ts';
import type {ColumnProperties, Worksheet} from '../../core/worksheet.ts';
import type {StyleRegistry} from './styles.ts';
import {
  buildColumnDefaults,
  buildPackageParts,
  createStyleRegistry,
  renderRow,
  type FlushedSheet,
  type WriteOptions,
} from './write.ts';

/** Calculation settings applied to the streamed workbook. Mirrors the {@link Workbook} flags. */
export interface CalcProperties {
  /** Ask the consumer to recalculate every formula on open — the OOXML `fullCalcOnLoad` flag. */
  fullCalcOnLoad?: boolean;
}

/** Options fixed at construction that shape the whole streamed package. */
export interface WorkbookStreamWriterOptions {
  /**
   * Pool plain string cell values into a shared-strings table rather than storing each inline — the
   * same {@link WriteOptions.useSharedStrings} the buffered writer exposes. Off by default.
   */
  readonly useSharedStrings?: boolean;

  /**
   * Write the package to a caller-owned {@link Writable} sink — an outbound upload, a cloud-SDK
   * stream, any destination the caller controls. {@link commit} pushes every chunk into it and
   * settles only once the sink has finished (or rejects if it errors), so a caller can deterministically
   * sequence work after the upload completes. Mutually exclusive with {@link filename}.
   */
  readonly stream?: Writable;

  /**
   * Write the package to a file at this path. The writer opens a `fs.createWriteStream` for it; if the
   * destination cannot be opened (bad path, name too long) the stream errors and {@link commit} rejects
   * with that I/O error rather than hanging. Mutually exclusive with {@link stream}.
   */
  readonly filename?: string;
}

/**
 * A row appended to a {@link WorksheetStreamWriter}. Style its cells through {@link cells}, then call
 * {@link commit} to mark it finished. In an eager (inline-strings) writer, committing serialises the
 * row and frees its cells from the model, bounding peak memory; with `useSharedStrings` on it is a
 * no-op and the row stays live until the workbook commits.
 */
export class StreamedRow {
  readonly #cells: readonly Cell[];
  readonly #sheet: WorksheetStreamWriter | null;
  readonly #number: number;
  #committed = false;

  constructor(cells: readonly Cell[], sheet: WorksheetStreamWriter | null, number: number) {
    this.#cells = cells;
    this.#sheet = sheet;
    this.#number = number;
  }

  /** The cells this row materialised, for styling before it is committed. */
  get cells(): readonly Cell[] {
    return this.#cells;
  }

  /** Finalise the row: an eager writer serialises it now and releases its cells; otherwise a no-op.
   * Committing twice is harmless — the second call does nothing rather than re-emitting the row. */
  commit(): void {
    if (this.#committed) return;
    this.#committed = true;
    this.#sheet?.flushRow(this.#number, this.#cells);
  }
}

/**
 * A worksheet being written incrementally. Append rows with {@link addRow}/{@link addRows}, style
 * cells through {@link getCell}, then {@link commit} to freeze it — after which any further mutation
 * is rejected with a legible error rather than silently accepted or crashing.
 */
export class WorksheetStreamWriter {
  readonly #sheet: Worksheet;
  readonly #eager: boolean;
  readonly #styles: StyleRegistry;
  #committed = false;
  // The last row number this writer appended, tracked independently of the model: an eager writer
  // evicts flushed rows, which lowers the model's used range, so leaning on it would reuse numbers.
  #lastRow = 0;
  // The column defaults an eagerly-rendered row inherits, frozen at the first flush so every flushed
  // row composes against the same columns even as later ones are defined.
  #columnDefaults: ReadonlyMap<number, ColumnProperties> | undefined;
  readonly #flushedRows: {number: number; xml: string}[] = [];
  #top = Infinity;
  #left = Infinity;
  #bottom = -Infinity;
  #right = -Infinity;

  constructor(sheet: Worksheet, eager: boolean, styles: StyleRegistry) {
    this.#sheet = sheet;
    this.#eager = eager;
    this.#styles = styles;
  }

  /** The sheet's name. */
  get name(): string {
    return this.#sheet.name;
  }

  /** The number of rows written so far — spans gaps and formatted-only rows, like the model, and
   * survives the eviction of eagerly-flushed rows. */
  get rowCount(): number {
    return Math.max(this.#lastRow, this.#sheet.rowCount);
  }

  /** Append one row of values after the last used row; the cells are returned for styling. */
  addRow(values: CellValue[]): StreamedRow {
    this.#assertOpen();
    if (!this.#eager) return new StreamedRow(this.#sheet.addRow(values), null, 0);
    const number = this.#nextRowNumber();
    return new StreamedRow(this.#placeRow(number, values), this, number);
  }

  /** Append a batch of rows in one call, each landing directly below the previous. */
  addRows(rows: CellValue[][]): StreamedRow[] {
    this.#assertOpen();
    if (!this.#eager) return this.#sheet.addRows(rows).map(cells => new StreamedRow(cells, null, 0));
    return rows.map(values => this.addRow(values));
  }

  // The next append position: past both this writer's own high-water mark and any rows a `getCell`
  // materialised, so appends never collide with random-access edits or with already-evicted rows.
  #nextRowNumber(): number {
    this.#lastRow = Math.max(this.#lastRow, this.#sheet.rowCount) + 1;
    return this.#lastRow;
  }

  // Materialise a positional row at an explicit number (the model's own append would reuse numbers once
  // eviction shrinks its used range), returning the cells for styling. A hole leaves its cell absent.
  #placeRow(number: number, values: CellValue[]): Cell[] {
    const cells: Cell[] = [];
    values.forEach((value, index) => {
      const cell = this.#sheet.getCell(encodeAddress(index + 1, number));
      cell.value = value;
      cells.push(cell);
    });
    return cells;
  }

  /**
   * Serialise an eagerly-committed row and release its cells from the model. Called by
   * {@link StreamedRow.commit}; the row's `<row>` XML is retained (interned into the workbook's live
   * style registry so its ids stay valid) and the cell graph is dropped, bounding peak memory.
   *
   * @throws {Error} if the row carries a shared-formula cell — a finished row cannot join the
   *   whole-sheet formula planning, so shared formulas must be authored through {@link getCell}.
   */
  flushRow(number: number, cells: readonly Cell[]): void {
    for (const cell of cells) {
      if (isSharedFormulaValue(cell.value)) {
        throw new Error(
          `row ${number} of streamed sheet "${this.#sheet.name}" carries a shared-formula cell; a ` +
            'committed row is finalised before the sheet is planned, so author shared formulas through ' +
            'getCell (leaving the row uncommitted) instead'
        );
      }
    }
    this.#columnDefaults ??= buildColumnDefaults(this.#sheet);
    const {xml, minCol, maxCol} = renderRow(
      {number, cells, properties: this.#sheet.rowProperties(number)},
      {
        columnDefaults: this.#columnDefaults,
        styles: this.#styles,
        sharedStrings: null,
        sharedRoles: new Map(),
        collapsedSummaries: new Set(),
      }
    );
    if (xml !== '') {
      this.#flushedRows.push({number, xml});
      if (minCol !== Infinity) {
        if (number < this.#top) this.#top = number;
        if (number > this.#bottom) this.#bottom = number;
        if (minCol < this.#left) this.#left = minCol;
        if (maxCol > this.#right) this.#right = maxCol;
      }
    }
    this.#sheet.evictRow(number);
  }

  // The rows this writer flushed, or undefined if none — handed to buildPackageParts at commit.
  flushedSheet(): FlushedSheet | undefined {
    if (this.#flushedRows.length === 0) return undefined;
    return {rows: this.#flushedRows, top: this.#top, left: this.#left, bottom: this.#bottom, right: this.#right};
  }

  /** Address a cell by its A1 reference to read or style it before the sheet is committed. */
  getCell(reference: string): Cell {
    this.#assertOpen();
    return this.#sheet.getCell(reference);
  }

  /**
   * Attach a data validation to a range before the sheet is committed. Delegates to the model, so the
   * streamed package emits the `<dataValidations>` block in its CT_Worksheet position — before
   * `<hyperlinks>` — because both writers share one worksheet serializer.
   */
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    this.#assertOpen();
    this.#sheet.addDataValidation(sqref, rule, options);
  }

  /**
   * Attach a conditional formatting to a range before the sheet is committed. Like every other block,
   * it lands in its schema-mandated slot — after `<mergeCells>`, before `<dataValidations>` and
   * `<hyperlinks>` — since the streamed sheet is serialized through the same path as a buffered write.
   */
  addConditionalFormatting(formatting: ConditionalFormatting): void {
    this.#assertOpen();
    this.#sheet.addConditionalFormatting(formatting);
  }

  /**
   * Anchor a workbook image (the id from {@link WorkbookStreamWriter.addImage}) to this sheet,
   * spanning the rectangle from the top-left grid point `tl` to the bottom-right `br`. The streamed
   * package emits the drawing part, its media relationship, and the sheet's `<drawing>` reference
   * exactly as a buffered write does — both writers share `buildPackageParts`.
   */
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void {
    this.#assertOpen();
    this.#sheet.addImage(imageId, anchor);
  }

  /**
   * Apply the sheet's autofilter before it is committed; mirrors {@link Worksheet.autoFilter}. The
   * streamed package emits `<autoFilter>` in its CT_Worksheet slot — after `<sheetProtection>` — and
   * contributes the hidden `_FilterDatabase` defined name, exactly as a buffered write does.
   */
  set autoFilter(filter: string | AutoFilter | undefined) {
    this.#assertOpen();
    this.#sheet.autoFilter = filter;
  }

  get autoFilter(): AutoFilter | undefined {
    return this.#sheet.autoFilter;
  }

  /**
   * Apply sheet-level protection before the sheet is committed; mirrors {@link Worksheet.protect}. The
   * shared serializer places `<sheetProtection>` ahead of `<autoFilter>` per CT_Worksheet, so a
   * streamed sheet carrying both stays valid rather than corrupt.
   */
  protect(password?: string, options: SheetProtectionOptions = {}): void {
    this.#assertOpen();
    this.#sheet.protect(password, options);
  }

  /** Freeze the sheet: no more rows or edits may be added after this. */
  commit(): void {
    this.#committed = true;
  }

  /** Whether the sheet has been committed. */
  get committed(): boolean {
    return this.#committed;
  }

  // Internal: the underlying model sheet, so the workbook writer can serialise it at commit time.
  get model(): Worksheet {
    return this.#sheet;
  }

  #assertOpen(): void {
    if (this.#committed) {
      throw new Error(
        `worksheet "${this.#sheet.name}" is already committed — its rows are finalised and no more can be added`
      );
    }
  }
}

/**
 * A workbook written incrementally to a Node stream. Add worksheets, append their rows, commit each
 * sheet, then {@link commit} the workbook to assemble and stream the package. The produced bytes are
 * available both as the resolved value of `commit()` and through {@link stream} (a Node `Readable`
 * that a caller can `pipe`).
 */
export class WorkbookStreamWriter {
  readonly #workbook = new Workbook();
  readonly #sheets: WorksheetStreamWriter[] = [];
  readonly #writeOptions: WriteOptions;
  readonly #sink: Writable | undefined;
  // The single style registry shared by the eager per-row flush and the commit-time serialisation, so a
  // flushed row's style ids match the styles.xml built from the same table.
  readonly #styles: StyleRegistry;
  // Eager per-row flushing runs with strings inline; a shared-strings pool is inherently whole-workbook,
  // so it cannot bound memory — turning it on keeps every row live until commit.
  readonly #eager: boolean;
  #stream: PassThrough | undefined;
  #committed = false;

  /** Calculation settings for the workbook; set `fullCalcOnLoad` before committing to emit it. */
  readonly calcProperties: CalcProperties = {};

  constructor(options: WorkbookStreamWriterOptions = {}) {
    this.#writeOptions = {useSharedStrings: options.useSharedStrings ?? false};
    this.#eager = !this.#writeOptions.useSharedStrings;
    this.#styles = createStyleRegistry(this.#workbook);
    if (options.stream && options.filename) {
      throw new Error('provide either a stream or a filename to the streaming writer, not both');
    }
    this.#sink = options.stream ?? (options.filename ? createWriteStream(options.filename) : undefined);
  }

  /** Document-level metadata written to the package's core properties. */
  get properties(): Workbook['properties'] {
    return this.#workbook.properties;
  }

  /**
   * The output stream carrying the package bytes. A caller drives it with Node's standard idiom —
   * `writer.stream.pipe(out)` — which composes because `pipe` returns its destination. The stream is
   * created lazily on first access so a caller handing the writer its own sink is still free to
   * ignore this one.
   */
  get stream(): Readable {
    return (this.#stream ??= new PassThrough());
  }

  /**
   * Register a picture's bytes on the workbook's shared media registry and return its id, to anchor
   * on any sheet with {@link WorksheetStreamWriter.addImage}. Mirrors {@link Workbook.addImage}: one
   * media part backs an image anchored on several sheets. Rejected once the workbook is committed.
   */
  addImage(options: AddImageOptions): number {
    if (this.#committed) {
      throw new Error('the workbook is already committed — no more images can be registered');
    }
    return this.#workbook.addImage(options);
  }

  /** Create a worksheet and append it to the workbook. */
  addWorksheet(name: string, options: AddWorksheetOptions = {}): WorksheetStreamWriter {
    if (this.#committed) {
      throw new Error('the workbook is already committed — no more worksheets can be added');
    }
    const sheet = new WorksheetStreamWriter(
      this.#workbook.addWorksheet(name, options),
      this.#eager,
      this.#styles
    );
    this.#sheets.push(sheet);
    return sheet;
  }

  /**
   * Assemble the workbook into its package, stream the bytes through {@link stream}, and resolve with
   * the same bytes. Every sheet is frozen first, so a row added after this rejects legibly. Idempotent
   * only in that a second call throws rather than re-emitting.
   */
  async commit(): Promise<Uint8Array> {
    if (this.#committed) {
      throw new Error('the workbook is already committed');
    }
    this.#committed = true;
    for (const sheet of this.#sheets) sheet.commit();
    if (this.calcProperties.fullCalcOnLoad) this.#workbook.fullCalcOnLoad = true;

    // Hand every sheet's eagerly-flushed rows to the shared serialiser, which emits them alongside the
    // rows still live in the model. The style registry is the same one the flushed rows interned into,
    // so styles.xml stays consistent with the ids already baked into their XML.
    const flushed = new Map<Worksheet, FlushedSheet>();
    for (const sheet of this.#sheets) {
      const sheetFlushed = sheet.flushedSheet();
      if (sheetFlushed) flushed.set(sheet.model, sheetFlushed);
    }
    const parts = buildPackageParts(this.#workbook, {
      ...this.#writeOptions,
      styles: this.#styles,
      flushed,
    });
    const owned = this.#stream;
    const sink = this.#sink;
    // Track the caller sink's terminal state before writing a byte, so an open failure that errors on a
    // later tick (a bad filename) is caught rather than lost — the whole point of the reject-not-hang
    // contract.
    const sinkSettled = sink ? settleOnFinish(sink) : undefined;

    const bytes = await streamZipPackage(parts, chunk => {
      owned?.write(chunk);
      sink?.write(chunk);
    });
    owned?.end();
    sink?.end();

    await sinkSettled;
    return bytes;
  }
}

// Resolve when a caller-supplied sink has flushed the whole package (`finish`), or reject if it errors
// (`error`) — the commit promise must settle either way, never hang. Whichever fires first wins; the
// other is ignored.
function settleOnFinish(sink: Writable): Promise<void> {
  return new Promise((resolve, reject) => {
    sink.once('finish', resolve);
    sink.once('error', reject);
  });
}

// Zip the package parts through fflate's streaming container, forwarding each output chunk to `onChunk`
// as it is produced and resolving with the whole archive once the final chunk arrives. `ZipDeflate`
// deflates synchronously, so the callback fires inline as each part is pushed — the CRC-32 fflate
// stamps into every entry's header therefore always matches the bytes it just compressed.
function streamZipPackage(
  parts: Record<string, Uint8Array>,
  onChunk: (chunk: Uint8Array) => void
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const collected: Uint8Array[] = [];
    const zip = new Zip((err, chunk, final) => {
      if (err) {
        reject(err);
        return;
      }
      collected.push(chunk);
      onChunk(chunk);
      if (final) resolve(concat(collected));
    });
    for (const [name, data] of Object.entries(parts)) {
      const entry = new ZipDeflate(name, {level: 6});
      zip.add(entry);
      entry.push(data, true);
    }
    zip.end();
  });
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
