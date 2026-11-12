// The streaming `.xlsx` writer: author a workbook incrementally and emit its package through a Node
// stream, rather than holding the finished bytes in one buffer as {@link writeXlsx} does.
//
// A producer adds worksheets, appends rows as it generates them, commits each sheet, then commits the
// workbook. The output goes to whichever destination the caller chose at construction: its own
// pull-based `stream` to pipe, a caller-owned `Writable` sink (an outbound upload), or a `filename` the
// writer opens. In every case `commit()` settles: it resolves once a supplied sink has flushed the
// whole package and rejects if that sink errors (an unopenable file), never hanging on a finish signal
// that will not come. The output rides a genuinely streamed zip container (fflate's
// `Zip`/`ZipDeflate`), which computes each entry's CRC-32 incrementally, so the archive is
// well-formed by construction, and the defect the upstream "streaming writer emits a corrupt zip"
// reports describe is structurally absent here.
//
// Peak memory (ADR 0046): with strings inline, the default, each sheet's part is written into its own
// zip entry as its rows commit. The entry opens at the sheet's first committed row with the part's
// head, every committed row is rendered, deflated in batches and handed to the sink, and the sheet's
// commit writes the rows still open and the part's tail. A producer that commits as it goes therefore
// holds only its uncommitted rows, whatever the sheet's length. The rows intern into the workbook's
// live style registry, the same one that emits `xl/styles.xml` at the end, so their style ids stay
// correct. Writing in order is what bounds memory, and it is also what the API now requires:
// committing a row commits every uncommitted row above it, `getCell` refuses a row at or above the
// last one written, the head (columns, view, sheet properties, outline depth) is frozen at the first
// write, and `<dimension>` is omitted. Turning `useSharedStrings` on keeps every row live until the
// workbook commits, because the pool's indices come from the whole workbook, and those sheets go
// through `buildPackageParts` like a buffered write.
//
// What `WorksheetStreamWriter` does *not* expose, since a hand-maintained subset of `Worksheet` grows
// its gaps silently: a verb added to the sheet model simply never appears here, and nothing mechanical
// would say so. An exhaustiveness proof would be the wrong artifact, because most of `Worksheet` must
// not be streamable, so the honest one is this list. Two classes are absent because a streamed sheet
// cannot answer them honestly. Structural edits (`spliceRows`, `insertRow`, `duplicateRow`, and the
// column verbs) re-anchor rows an eager writer has already written and evicted, so they would apply
// to the live tail alone while looking like they applied to the sheet. Whole-sheet reads (`usedRange`,
// `getRange`, `hasCell`, `actualRowCount`, `dataValidationAt`) have the same defect one step later:
// each would compute a plausible answer over whichever rows happen to still be in memory. `rowCount`
// is the exception that proves it, and is why it keeps its own high-water mark instead of asking the
// model.
//
// The third class is different, and calling it absent-by-design would be a fiction: merges, the frozen
// pane, column widths, page setup, tables and comment threads are sheet-level state the streamed
// package already emits, because the part's head and tail are rendered by the same serialiser the
// buffered writer uses. They are missing because no streaming producer has asked for them yet. Adding
// one is a delegation and a test, subject to the ordering rule above: what the head holds is read at
// the first write, and a change after that is refused at the sheet's commit.

import {createWriteStream} from 'node:fs';
import {PassThrough, type Readable, type Writable} from 'node:stream';

import {encodeAddress, tryDecodeCellRef} from '../../core/address.ts';
import {type FormulaPlacement, formulaPlacement} from '../../core/array-formula-ranges.ts';
import type {AutoFilter} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import type {ConditionalFormatting} from '../../core/conditional-formatting.ts';
import type {DataValidation} from '../../core/data-validation.ts';
import type {DateEpoch} from '../../core/date.ts';
import {formulaNamesInScope} from '../../core/formula.ts';
import type {Hyperlink} from '../../core/hyperlink.ts';
import type {AnchorPoint} from '../../core/image.ts';
import {INTERNAL} from '../../core/internal.ts';
import type {SheetProtectionOptions} from '../../core/protection.ts';
import type {Row} from '../../core/row.ts';
import {type CellValue, isSharedFormulaValue} from '../../core/value.ts';
import {type AddImageOptions, type AddWorksheetOptions, Workbook} from '../../core/workbook.ts';
import type {ColumnProperties, Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, quoted} from '../../errors.ts';
import {WorkbookMetadataTable} from './cell-metadata.ts';
import {type CommentCell, collectNotes} from './comments.ts';
import {type PackageEntry, PackageStream} from './package-stream.ts';
import {worksheetPart} from './part-names.ts';
import {assertWritableLevel, buildColumnDefaults, renderRow} from './row-xml.ts';
import {planSharedFormulas, type SharedFormulaRole} from './shared-formulas.ts';
import {type OutlineRow, StreamedOutline} from './streamed-outline.ts';
import type {StyleRegistry} from './styles.ts';
import {
  assertNoArrayRangeConflict,
  validateMerges,
  worksheetHeadXml,
  type WorksheetHeadFacts,
  worksheetTailXml,
} from './worksheet-xml.ts';
import {
  type StreamedSheet,
  buildPackageParts,
  createStyleRegistry,
  planStreamedSheetTail,
  type WriteOptions,
} from './write.ts';

/** Calculation settings applied to the streamed workbook. Mirrors the {@link Workbook} flags. */
export interface CalcProperties {
  /** Ask the consumer to recalculate every formula on open: the OOXML `fullCalcOnLoad` flag. */
  fullCalcOnLoad?: boolean;
}

// A caller may supply a sink to write to (a stream or a filename) but never both. This is enforced as a
// type-level union rather than only the runtime check the constructor also carries, so passing both is
// a compile error, not a first-run surprise.
type SinkOptions =
  | {
      /**
       * Write the package to a caller-owned {@link Writable} sink: an outbound upload, a cloud-SDK
       * stream, any destination the caller controls. {@link WorkbookStreamWriter.commit} pushes every
       * chunk into it and settles only once the sink has finished (or rejects if it errors), so a
       * caller can deterministically sequence work after the upload completes. Mutually exclusive
       * with {@link filename}.
       */
      readonly stream: Writable;
      readonly filename?: never;
    }
  | {
      readonly stream?: never;
      /**
       * Write the package to a file at this path. The writer opens a `fs.createWriteStream` for it; if
       * the destination cannot be opened (bad path, name too long) the stream errors and
       * {@link WorkbookStreamWriter.commit} rejects with that I/O error rather than hanging. Mutually
       * exclusive with {@link stream}.
       */
      readonly filename: string;
    }
  | {readonly stream?: never; readonly filename?: never};

/** Options fixed at construction that shape the whole streamed package. */
export type WorkbookStreamWriterOptions = SinkOptions & {
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

/** Options for {@link WorkbookStreamWriter.addWorksheet}. */
export interface AddStreamedWorksheetOptions extends AddWorksheetOptions {
  /**
   * The deepest row outline level the sheet will use, written as `<sheetFormatPr outlineLevelRow>`.
   * That element precedes the rows, so a streamed sheet writes it when its first row is committed,
   * from the larger of this and the deepest level among the rows written then. A row committed later
   * with a deeper level is refused, so declare the depth here when the outline starts further down
   * the sheet. A consumer sizes the outline bar from it.
   */
  readonly outlineLevelRow?: number;
}

/**
 * A row appended to a {@link WorksheetStreamWriter}. Style its cells through {@link cells}, then call
 * {@link commit} to mark it finished. In an eager (inline-strings) writer, committing writes the row
 * into the sheet's zip entry and frees its cells; with `useSharedStrings` on it is a no-op and the row
 * stays live until the workbook commits.
 */
export class StreamedRow {
  readonly #cells: readonly Cell[];
  readonly #sheet: WorksheetStreamWriter | null;
  readonly #number: number;

  // Private, and reached from `WorksheetStreamWriter.addRow` through the static channel below. A
  // public constructor let `new StreamedRow([], sheetWriter, n).commit()` evict row `n` from the model
  // and write nothing for it: the number is the writer's to hand out, not a caller's to name.
  private constructor(cells: readonly Cell[], sheet: WorksheetStreamWriter | null, number: number) {
    this.#cells = cells;
    this.#sheet = sheet;
    this.#number = number;
  }

  /** The sheet writer's construction channel, keyed like its own; see `core/internal.ts`. */
  static readonly [INTERNAL]: StreamedRowFactory = {
    create(cells, sheet, number) {
      return new StreamedRow(cells, sheet, number);
    },
  };

  /** The cells `addRow` materialised, for styling before it is committed. A cell added to the row
   * through {@link WorksheetStreamWriter.getCell} is not among them, and is committed all the same. */
  get cells(): readonly Cell[] {
    return this.#cells;
  }

  /**
   * Finalise the row. An eager writer writes it, and every uncommitted row above it, into the sheet's
   * zip entry and releases their cells; otherwise this is a no-op. Committing twice, or committing a
   * row a later row's commit already wrote, does nothing.
   *
   * The promise resolves once the caller's sink can take more, so a producer that awaits it is paced
   * by the sink. It never rejects: a sink that fails is reported by {@link WorkbookStreamWriter.commit}.
   *
   * @throws AuthoringError if a row being written carries a shared-formula cell, a row outline level
   *   deeper than the sheet's head declares, or an array formula whose range collides with another
   *   formula. See {@link WorksheetStreamWriterInternals.commitThrough}.
   */
  commit(): Promise<void> {
    return this.#sheet?.[INTERNAL].commitThrough(this.#number) ?? Promise.resolve();
  }
}

/**
 * How a {@link StreamedRow} comes into being. A row belongs to the sheet writer that numbered it, and
 * only that writer knows which number is still open. `sheet` is `null` for a writer that keeps every
 * row live until the workbook commits, where committing a row has nothing to do. Reached as
 * `StreamedRow[INTERNAL]`.
 */
export interface StreamedRowFactory {
  create(cells: readonly Cell[], sheet: WorksheetStreamWriter | null, number: number): StreamedRow;
}

/**
 * The workbook writer's plumbing a sheet writer is built with. A consumer supplies none of it: the
 * interning table its rows' style ids come from, the epoch its dates count in, the archive its part
 * is written into.
 */
export interface WorksheetStreamContext {
  readonly workbook: Workbook;
  readonly sheet: Worksheet;
  /** The sheet's position in the workbook, which names its part. */
  readonly index: number;
  readonly eager: boolean;
  readonly styles: StyleRegistry;
  readonly cellMetadata: WorkbookMetadataTable;
  readonly output: PackageStream;
  readonly outlineLevelRow: number;
}

// What the head was written with, kept to check at sheet commit that nothing it holds changed since.
interface WrittenHead {
  readonly xml: string;
  readonly facts: WorksheetHeadFacts;
}

/**
 * A worksheet being written incrementally. Append rows with {@link addRow}/{@link addRows}, style
 * cells through {@link getCell}, then {@link commit} to finish it, after which any further mutation
 * is rejected with a legible error rather than silently accepted or crashing.
 */
export class WorksheetStreamWriter {
  readonly #context: WorksheetStreamContext;
  readonly #sheet: Worksheet;
  #committed: Promise<void> | undefined;
  // The last row number this writer appended, tracked independently of the model: an eager writer
  // evicts written rows, which lowers the model's used range, so leaning on it would reuse numbers.
  #lastRow = 0;
  // The highest row number written into the entry. Rows are written in ascending order, so nothing at
  // or above it can be written any more.
  #written = 0;
  #entry: PackageEntry | undefined;
  #head: WrittenHead | undefined;
  // The column defaults a written row inherits, frozen with the head so every row composes against
  // the columns `<cols>` declares.
  #columnDefaults: ReadonlyMap<number, ColumnProperties> = new Map();
  // Built at the first write, with the head, because it depends on the outline placement the head
  // records.
  #outline: StreamedOutline | undefined;
  // What a written row carried that is serialised outside its `<row>`, into the comments and VML
  // parts, taken before eviction.
  readonly #notes: CommentCell[] = [];
  // The array formulas written so far whose range reaches past the last row written: the only ones a
  // row still to come can collide with.
  #openArrays: FormulaPlacement[] = [];
  #record: StreamedSheet | undefined;

  // Private, and reached from `WorkbookStreamWriter.addWorksheet` through the static channel below.
  // A caller never builds one of these -- they receive it from `addWorksheet` -- and the context says
  // why that matters: a `StyleRegistry` is the writer's own interning table, so a public constructor
  // made it a type a consumer could hold and could not name.
  private constructor(context: WorksheetStreamContext) {
    this.#context = context;
    this.#sheet = context.sheet;
  }

  /**
   * The workbook writer's construction channel. A static member keyed by the same symbol as the
   * instance one below, so both halves of this class's internal wiring are reached the one way; see
   * `core/internal.ts`.
   */
  static readonly [INTERNAL]: WorksheetStreamWriterFactory = {
    create(context) {
      return new WorksheetStreamWriter(context);
    },
  };

  /** The sheet's name. */
  get name(): string {
    return this.#sheet.name;
  }

  /** The number of rows written so far. Spans gaps and formatted-only rows, like the model, and
   * survives the eviction of written rows. */
  get rowCount(): number {
    return Math.max(this.#lastRow, this.#written, this.#sheet.rowCount);
  }

  /** Append one row of values after the last used row; the cells are returned for styling. */
  addRow(values: CellValue[]): StreamedRow {
    this.#assertOpen();
    if (!this.#context.eager) {
      return StreamedRow[INTERNAL].create(this.#sheet.addRow(values), null, 0);
    }
    const number = this.#nextRowNumber();
    return StreamedRow[INTERNAL].create(this.#placeRow(number, values), this, number);
  }

  /** Append a batch of rows in one call, each landing directly below the previous. */
  addRows(rows: CellValue[][]): StreamedRow[] {
    this.#assertOpen();
    if (!this.#context.eager)
      return this.#sheet.addRows(rows).map((cells) => StreamedRow[INTERNAL].create(cells, null, 0));
    return rows.map((values) => this.addRow(values));
  }

  // The next append position: past this writer's own high-water mark, any rows a `getCell`
  // materialised, and the rows already written, so appends never collide with any of them.
  #nextRowNumber(): number {
    this.#lastRow = this.rowCount + 1;
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

  // Write every live row numbered up to `number`, in order. The rows are read off the model as it
  // stands now rather than as `addRow` left them: `getCell` may add cells to an open row, and the
  // eviction takes every cell the model holds for it, so a cell missing from what is rendered here
  // would be gone from the file with nothing to say so.
  #commitThrough(number: number): Promise<void> {
    if (number <= this.#written) return this.#context.output.ready();
    this.#assertOpen();
    const rows = [...this.#sheet.rows()].filter((row) => row.number <= number);
    for (const row of rows) {
      for (const cell of row.cells) {
        if (isSharedFormulaValue(cell.value)) {
          throw new AuthoringError(
            `row ${row.number} of streamed sheet ${quoted(this.#sheet.name)} carries a shared-formula cell; ` +
              'a committed row is written before the rows below it exist, so author shared formulas ' +
              'through getCell and leave their rows uncommitted until the sheet commits',
          );
        }
      }
    }
    this.#writeRows(rows, new Map());
    this.#written = number;
    return this.#context.output.ready();
  }

  // Write `rows` (ascending, all below the last row written) into the entry, opening it with the head
  // first if this is the first write, then evict them. Everything that can refuse the batch is asked
  // before anything is written, so a refused commit leaves the entry as it was.
  #writeRows(rows: readonly Row[], sharedRoles: ReadonlyMap<string, SharedFormulaRole>): void {
    const outlineRows: OutlineRow[] = rows.map(({number, properties}) => ({
      number,
      level: properties?.outlineLevel ?? 0,
      hidden: properties?.hidden ?? false,
    }));
    const head = this.#head ?? this.#renderHead(outlineRows);
    for (const {number, level} of outlineRows) {
      assertWritableLevel(`row ${number} outlineLevel`, level);
      if (level > head.facts.outlineLevelRow) {
        throw new AuthoringError(
          `row ${number} of streamed sheet ${quoted(this.#sheet.name)} has outline level ${level}, ` +
            `deeper than the ${head.facts.outlineLevelRow} its head was written with when its first ` +
            'row was committed; declare the depth with the outlineLevelRow option of addWorksheet',
        );
      }
    }
    const placements: FormulaPlacement[] = [];
    for (const {cells} of rows) {
      for (const cell of cells) {
        const placement = formulaPlacement(cell.address, cell.col, cell.row, cell.value);
        if (placement !== undefined) placements.push(placement);
      }
    }
    assertNoArrayRangeConflict([...this.#openArrays, ...placements]);

    this.#outline ??= new StreamedOutline(this.#sheet.outline.summaryBelow !== false);
    const collapsedSummaries = this.#outline.collapsedIn(outlineRows);
    // The first write reads the column defaults with the rest of the head; every later row composes
    // against the same ones, because `<cols>` is already written.
    this.#columnDefaults =
      this.#head === undefined ? buildColumnDefaults(this.#sheet) : this.#columnDefaults;
    const context = {
      columnDefaults: this.#columnDefaults,
      styles: this.#context.styles,
      sharedStrings: null,
      sharedRoles,
      collapsedSummaries,
      dateEpoch: this.#context.workbook.dateEpoch,
      formulaNames: formulaNamesInScope(this.#context.workbook.definedNames, this.#sheet.name),
      cellMetadata: this.#context.cellMetadata,
    };
    const rendered = rows.map((row) => renderRow(row, context).xml);

    const entry = this.#open(head);
    for (const xml of rendered) entry.write(xml);
    for (const row of rows) {
      this.#notes.push(...collectNotes(row.cells));
      this.#sheet[INTERNAL].evictRow(row.number);
    }
    const last = Math.max(rows.at(-1)?.number ?? 0, this.#written);
    this.#written = last;
    this.#openArrays = [...this.#openArrays, ...placements].filter(
      (placement) => placement.arrayRef !== undefined && arrayBottom(placement) > last,
    );
  }

  // The head as it would be written now. The first write takes the deepest level among its own rows
  // as well as the declared one, so a sheet whose outline is in its first batch needs no declaration.
  #renderHead(firstRows: readonly OutlineRow[]): WrittenHead {
    let outlineLevelRow = this.#context.outlineLevelRow;
    for (const {level} of firstRows) outlineLevelRow = Math.max(outlineLevelRow, level);
    const facts = {dimensionRef: null, outlineLevelRow};
    return {xml: this.#headXml(facts), facts};
  }

  #headXml(facts: WorksheetHeadFacts): string {
    const {workbook, index, styles} = this.#context;
    return worksheetHeadXml(this.#sheet, styles, index === workbook.activeTabIndex, facts);
  }

  #open(head: WrittenHead): PackageEntry {
    if (this.#entry !== undefined) return this.#entry;
    this.#head = head;
    this.#entry = this.#context.output.open(worksheetPart(this.#context.index + 1));
    this.#entry.write(`${head.xml}<sheetData>`);
    return this.#entry;
  }

  // Write the rows still live and the part's tail, and close the entry.
  #finish(): void {
    if (!this.#context.eager) return;
    // The rows never committed were never written, so they can still join the whole-sheet shared
    // formula plan: every row the model holds now is one of them.
    this.#writeRows([...this.#sheet.rows()], planSharedFormulas(this.#sheet));
    const head = this.#head;
    const entry = this.#entry;
    if (head === undefined || entry === undefined) return;
    if (this.#headXml(head.facts) !== head.xml) {
      throw new AuthoringError(
        `the columns, view, outline placement or properties of streamed sheet ${quoted(this.#sheet.name)} ` +
          'changed after its first row was written, and they precede the rows in the file; set them ' +
          'before the first row is committed',
      );
    }
    validateMerges(this.#sheet);
    const {workbook, index, styles} = this.#context;
    const tail = planStreamedSheetTail(workbook, index, this.#notes);
    entry.write(`</sheetData>${worksheetTailXml({sheet: this.#sheet, styles, ...tail})}`);
    entry.close();
    this.#record = {notes: this.#notes, wiring: tail.wiring};
  }

  /**
   * Address a cell by its A1 reference to read or style it before its row is written. Rows are written
   * in order, so the row must be below the last one written.
   *
   * @throws AuthoringError if the reference names a row at or above the last row written. That row,
   *   and every row above it, is already in the file and its cells are released, so the cell this
   *   would materialise could only be written as a second row out of order.
   */
  getCell(reference: string): Cell {
    this.#assertOpen();
    const target = tryDecodeCellRef(reference);
    if (target !== undefined && target.row <= this.#written) {
      throw new AuthoringError(
        `cell ${reference} of streamed sheet ${quoted(this.#sheet.name)} is in row ${target.row}, but ` +
          `rows up to ${this.#written} are already written: a streamed sheet is written in order, ` +
          'and its committed rows are released. Fill a row before committing it or any row below it.',
      );
    }
    return this.#sheet.getCell(reference);
  }

  /**
   * Attach a data validation to a range before the sheet is committed. Delegates to the model, so the
   * streamed package emits the `<dataValidations>` block in its CT_Worksheet position, before
   * `<hyperlinks>`, because both writers share one worksheet serializer.
   */
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    this.#assertOpen();
    this.#sheet.addDataValidation(sqref, rule, options);
  }

  /**
   * Put a hyperlink on a cell or a rectangle of cells before the sheet is committed; mirrors
   * {@link Worksheet.addHyperlink}. A link lives beside the grid rather than in a row, so it may cover
   * cells of a row that is already committed.
   */
  addHyperlink(link: Hyperlink): void {
    this.#assertOpen();
    this.#sheet.addHyperlink(link);
  }

  /**
   * Attach a conditional formatting to a range before the sheet is committed. Like every other block,
   * it lands in its schema-mandated slot, after `<mergeCells>` and before `<dataValidations>` and
   * `<hyperlinks>`, since the streamed sheet is serialized through the same path as a buffered write.
   */
  addConditionalFormatting(formatting: ConditionalFormatting): void {
    this.#assertOpen();
    this.#sheet.addConditionalFormatting(formatting);
  }

  /**
   * Anchor a workbook image (the id from {@link WorkbookStreamWriter.addImage}) to this sheet,
   * spanning the rectangle from the top-left grid point `tl` to the bottom-right `br`. The streamed
   * package emits the drawing part, its media relationship, and the sheet's `<drawing>` reference
   * exactly as a buffered write does.
   */
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void {
    this.#assertOpen();
    this.#sheet.addImage(imageId, anchor);
  }

  /**
   * Apply the sheet's autofilter before it is committed; mirrors {@link Worksheet.autoFilter}. The
   * streamed package emits `<autoFilter>` in its CT_Worksheet slot, after `<sheetProtection>`, and
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

  /**
   * Finish the sheet: an eager writer writes the rows still uncommitted and the rest of the part, and
   * closes its zip entry. No more rows or edits may be added after this; committing again returns the
   * same promise. The promise resolves as {@link StreamedRow.commit}'s does.
   *
   * @throws AuthoringError if a row still to be written is refused (see {@link StreamedRow.commit}),
   *   if a column, the view or a sheet property changed after the first row was written, or if the
   *   sheet's merges overlap a table.
   */
  commit(): Promise<void> {
    if (this.#committed !== undefined) return this.#committed;
    this.#finish();
    this.#committed = this.#context.output.ready();
    return this.#committed;
  }

  /** Whether the sheet has been committed. */
  get committed(): boolean {
    return this.#committed !== undefined;
  }

  /**
   * The underlying model sheet: row properties (an outline level, a height), column definitions and the
   * view are set here. What precedes the rows in the file (columns, view, sheet properties) is read
   * when the first row is written; set it before then.
   */
  get model(): Worksheet {
    return this.#sheet;
  }

  readonly [INTERNAL]: WorksheetStreamWriterInternals = {
    commitThrough: (number) => this.#commitThrough(number),
    streamed: () => this.#record,
  };

  #assertOpen(): void {
    if (this.#committed !== undefined) {
      throw new AuthoringError(
        `worksheet ${quoted(this.#sheet.name)} is already committed: its rows are finalised and no more can be added`,
      );
    }
  }
}

// The last row an array formula's range reaches, or its own row when the range is unreadable (the
// writer refuses such a formula elsewhere).
function arrayBottom(placement: FormulaPlacement): number {
  const match = /(\d+)$/.exec(placement.arrayRef ?? '');
  return match?.[1] === undefined ? placement.row : Number(match[1]);
}

/**
 * How a {@link WorksheetStreamWriter} comes into being. Reached as `WorksheetStreamWriter[INTERNAL]`;
 * a consumer receives the sheet writer from {@link WorkbookStreamWriter.addWorksheet}.
 */
export interface WorksheetStreamWriterFactory {
  create(context: WorksheetStreamContext): WorksheetStreamWriter;
}

/**
 * What the streaming writer's own machinery may do to a sheet writer that a caller may not, the same
 * boundary `sheet[INTERNAL]` draws around the model. Reached as `writer[INTERNAL]`; see
 * `core/internal.ts`.
 */
export interface WorksheetStreamWriterInternals {
  /**
   * Write every uncommitted row numbered up to `number` into the sheet's zip entry, in order, and
   * release their cells. Called by {@link StreamedRow.commit}; a no-op for a row already written.
   *
   * @throws AuthoringError if one of those rows carries a shared-formula cell (a committed row cannot
   *   join the whole-sheet formula plan, so author shared formulas through
   *   {@link WorksheetStreamWriter.getCell} and leave their rows to the sheet's commit), has an outline
   *   level deeper than the head declares, or holds an array formula whose range collides with
   *   another formula.
   */
  commitThrough(number: number): Promise<void>;

  /** What was recorded about the sheet once its entry is closed, or undefined before that. */
  streamed(): StreamedSheet | undefined;
}

/**
 * A workbook written incrementally to a Node stream. Add worksheets, append and commit their rows,
 * commit each sheet, then {@link commit} the workbook to write the rest of the package. With strings
 * inline, the default, a sheet's rows reach the destination as they are committed (ADR 0046).
 */
export class WorkbookStreamWriter {
  readonly #workbook = new Workbook();
  readonly #sheets: WorksheetStreamWriter[] = [];
  readonly #writeOptions: WriteOptions;
  readonly #sink: Writable | undefined;
  readonly #output: PackageStream;
  // The single style registry shared by the written rows and the commit-time serialisation, so a
  // written row's style ids match the styles.xml built from the same table.
  readonly #styles: StyleRegistry;
  // Shared the same way: a written row's dynamic-array formula or rich-value error points into the
  // parts this decides on.
  readonly #cellMetadata = new WorkbookMetadataTable();
  // Rows are written as they commit only with strings inline: a shared-strings pool is inherently
  // whole-workbook, so turning it on keeps every row live until commit.
  readonly #eager: boolean;
  #stream: PassThrough | undefined;
  #committed = false;

  /** Calculation settings for the workbook; set `fullCalcOnLoad` before committing to emit it. */
  readonly calcProperties: CalcProperties = {};

  constructor(options: WorkbookStreamWriterOptions = {}) {
    this.#writeOptions = {useSharedStrings: options.useSharedStrings ?? false};
    this.#eager = !this.#writeOptions.useSharedStrings;
    if (options.dateEpoch !== undefined) this.#workbook.dateEpoch = options.dateEpoch;
    this.#styles = createStyleRegistry(this.#workbook);
    if (options.stream && options.filename) {
      throw new AuthoringError(
        'provide either a stream or a filename to the streaming writer, not both',
      );
    }
    this.#sink =
      options.stream ?? (options.filename ? createWriteStream(options.filename) : undefined);
    // Built now, not at the first byte: it watches the sink for an error from the moment there is a
    // sink, so a file that fails to open is reported by `commit()` rather than crashing the process
    // as an `error` event nobody listened to.
    this.#output = new PackageStream(this.#sink);
  }

  /** Document-level metadata written to the package's core properties. */
  get properties(): Workbook['properties'] {
    return this.#workbook.properties;
  }

  /**
   * The output stream carrying the package bytes. A caller drives it with Node's standard idiom,
   * `writer.stream.pipe(out)`, which composes because `pipe` returns its destination. Created lazily
   * on first access, and from then on {@link commit} also resolves with the whole package, since this
   * stream may never be drained. It is not paced: the writer only waits on a sink it was given.
   *
   * @throws AuthoringError if the writer has a sink and has already written bytes to it: those bytes
   *   were never kept, so the stream would begin part-way through the archive. Take it before the
   *   first row is committed.
   */
  get stream(): Readable {
    if (this.#stream !== undefined) return this.#stream;
    if (this.#sink !== undefined && this.#output.started) {
      throw new AuthoringError(
        'the streaming writer has already written to its sink, so writer.stream would start part-way ' +
          'through the package; take writer.stream before the first row is committed',
      );
    }
    this.#stream = new PassThrough();
    this.#output.attach(this.#stream);
    return this.#stream;
  }

  /**
   * Register a picture's bytes on the workbook's shared media registry and return its id, to anchor
   * on any sheet with {@link WorksheetStreamWriter.addImage}. Mirrors {@link Workbook.addImage}: one
   * media part backs an image anchored on several sheets. Rejected once the workbook is committed.
   */
  addImage(options: AddImageOptions): number {
    this.#assertOpen('no more images can be registered');
    return this.#workbook.addImage(options);
  }

  /**
   * Create a worksheet and append it to the workbook. Its part is written into the archive as its rows
   * commit, so to keep memory bounded commit each sheet before writing the next: the archive holds one
   * entry at a time, and a later sheet's bytes wait in memory until the earlier one closes.
   */
  addWorksheet(name: string, options: AddStreamedWorksheetOptions = {}): WorksheetStreamWriter {
    this.#assertOpen('no more worksheets can be added');
    const {outlineLevelRow = 0, ...sheetOptions} = options;
    assertWritableLevel('outlineLevelRow', outlineLevelRow);
    const index = this.#workbook.worksheets.length;
    const sheet = WorksheetStreamWriter[INTERNAL].create({
      workbook: this.#workbook,
      sheet: this.#workbook.addWorksheet(name, sheetOptions),
      index,
      eager: this.#eager,
      styles: this.#styles,
      cellMetadata: this.#cellMetadata,
      output: this.#output,
      outlineLevelRow,
    });
    this.#sheets.push(sheet);
    return sheet;
  }

  /**
   * Commit every sheet still open, write the rest of the package, and end the destinations. A second
   * call throws rather than re-emitting.
   *
   * **What it resolves with, and why it can be `undefined`.** The archive is handed back only when
   * somebody asked for it: when no sink was supplied, or when {@link stream} was touched, since a
   * `PassThrough` nobody drained would otherwise be the only copy. A caller who passed `{filename}`
   * or their own `stream` and never reached for {@link stream} gets `undefined`, and the package is
   * never held as a whole: that is the memory the caller passed a sink to avoid. The type says so
   * rather than the prose alone, because a promise that resolves with a value the caller was told to
   * ignore is a promise they will use.
   *
   * If assembling the package fails, or the sink failed at any point since the first row was written,
   * the returned promise rejects *and* every stream this writer was given or handed out is destroyed
   * with that error. A caller piping {@link stream}, or awaiting its own sink, therefore sees an
   * `error` rather than waiting on a stream that will never end; and it is an `error` rather than a
   * clean `end` because the bytes written so far are a truncated package that must not be read as a
   * whole one.
   */
  async commit(): Promise<Uint8Array | undefined> {
    this.#assertOpen('its package is already assembled and cannot be re-emitted');
    this.#committed = true;
    try {
      for (const sheet of this.#sheets) await sheet.commit();
      if (this.calcProperties.fullCalcOnLoad) this.#workbook.fullCalcOnLoad = true;

      const streamed = new Map<Worksheet, StreamedSheet>();
      for (const sheet of this.#sheets) {
        const record = sheet[INTERNAL].streamed();
        if (record !== undefined) streamed.set(sheet.model, record);
      }
      // The style registry is the same one the written rows interned into, so styles.xml stays
      // consistent with the ids already in their XML.
      const parts = buildPackageParts(this.#workbook, {
        ...this.#writeOptions,
        styles: this.#styles,
        cellMetadata: this.#cellMetadata,
        streamed,
      });
      for (const [name, data] of Object.entries(parts)) await this.#output.writePart(name, data);
      return await this.#output.finish();
    } catch (error) {
      const failure = error instanceof Error ? error : new Error(String(error));
      this.#output.abort(failure);
      throw error;
    }
  }

  #assertOpen(what: string): void {
    if (this.#committed) {
      throw new AuthoringError(`the workbook is already committed: ${what}`);
    }
  }
}
