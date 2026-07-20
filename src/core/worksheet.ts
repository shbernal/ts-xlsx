// A worksheet: a sparsely-populated grid of cells addressed by A1 reference.
//
// Storage is sparse by construction — a spreadsheet is mostly empty, so cells
// materialise on first access and only occupied positions cost memory. Column and
// row metadata (widths, heights, visibility, outline grouping) are stored apart from
// the cell grid, because a column or row can carry formatting while holding no cells.
// Merges and views layer on in later slices.

import {decodeAddress, decodeRange, encodeAddress} from './address.ts';
import {type AutoFilter, canonicalizeAutoFilter} from './autofilter.ts';
import {applyCellStyle, Cell, cellToModel, copyCellContent} from './cell.ts';
import {type ConditionalFormatting, cloneConditionalFormatting} from './conditional-formatting.ts';
import {overwrite, replaceContents} from './containers.ts';
import {
  cloneDataValidation,
  type DataValidation,
  type DataValidationEntry,
} from './data-validation.ts';
import {GridEdits} from './grid-edits.ts';
import {
  type AnchoredImage,
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  PX_TO_EMU,
  type TwoCellAnchor,
} from './image.ts';
import {decodeSqrefRects, type MergeRect, rectsOverlap} from './merge.ts';
import type {HeaderFooter, PageBreak, PageMargins, PageSetup, PrintOptions} from './page-setup.ts';
import {type ParsedPivotTable, PivotTable, type PivotTableOptions} from './pivot-table.ts';
import type {PreservedWorksheetReference} from './preserved.ts';
import {
  deriveCredential,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionOptions,
} from './protection.ts';
import type {CellStyle, Color, Fill} from './style.ts';
import {Table, type TableOptions} from './table.ts';
import type {CellValue} from './value.ts';

export interface WorksheetState {
  /** Sheet visibility, as Excel models it. Defaults to `visible`. */
  readonly state: 'visible' | 'hidden' | 'veryHidden';
}

/** Format defaults applied to every row/column that carries no explicit override. */
export interface WorksheetProperties {
  /** Height, in points, for rows with no explicit height. */
  defaultRowHeight?: number;
  /** Width, in character units, for columns with no explicit width. */
  defaultColWidth?: number;
}

/**
 * Placement of an outline's summary rows/columns. Excel's defaults are summary *below* the detail
 * rows and *right* of the detail columns; setting either to `false` inverts that placement so an
 * author who groups upward gets a file that honours it. An unset flag is omitted from the written
 * `<outlinePr>`, and an empty object emits no `<outlinePr>` at all.
 */
export interface OutlineProperties {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}

/**
 * A worksheet's frozen-pane view. `state` `'frozen'` locks the top `ySplit` rows and left `xSplit`
 * columns in place while the rest scrolls; `'normal'` (the default) has no split and emits no
 * `<pane>` — writing a normal view leaves no leftover pane markup that would trip Excel's repair
 * prompt. An empty object is a normal view.
 */
export interface SheetView {
  /** Freeze state. Absent or `'normal'` means no split. */
  state?: 'normal' | 'frozen';
  /** Number of columns frozen at the left; `0`/absent freezes no columns. */
  xSplit?: number;
  /** Number of rows frozen at the top; `0`/absent freezes no rows. */
  ySplit?: number;
  /** The cell anchoring the bottom-right scrolling pane; defaults to the first unfrozen cell. */
  topLeftCell?: string;
}

/**
 * Per-column formatting. A column may exist purely to carry these, with no cells. The style
 * facets are *defaults* for the column's cells: a cell that sets a facet of its own wins, but
 * one that leaves a facet unset inherits the column's — the same precedence Excel applies, and
 * symmetric with how a {@link RowProperties} fill defaults a row's cells.
 */
export interface ColumnProperties extends CellStyle {
  /** Stable key naming the column so a keyed-object row (see {@link Worksheet.addRow}) can place a
   * value under it by name rather than position. In-memory only — it is not serialized to OOXML. */
  key?: string;
  /** Column width in character units. */
  width?: number;
  /** Whether the column is hidden. */
  hidden?: boolean;
  /** Outline (grouping) depth; 0 or absent means ungrouped. */
  outlineLevel?: number;
  /** Whether this column is the collapsed summary of an outline group. */
  collapsed?: boolean;
}

/** A row handed to {@link Worksheet.addRow}: a positional array of cell values (a hole or `undefined`
 * leaves that column untouched), or an object keyed by column {@link ColumnProperties.key} whose
 * values land under the matching columns. */
export type RowInput = (CellValue | undefined)[] | Record<string, CellValue>;

/** Per-row formatting. A row may exist purely to carry these, with no cells. */
export interface RowProperties {
  /** Row height in points. */
  height?: number;
  /** Whether the row is hidden. */
  hidden?: boolean;
  /** Outline (grouping) depth; 0 or absent means ungrouped. */
  outlineLevel?: number;
  /** Whether this row is the collapsed summary of an outline group. */
  collapsed?: boolean;
  /** Background fill applied to the row's cells that carry no fill of their own. */
  fill?: Fill;
}

/** One materialised cell in a {@link WorksheetModel}: its position, value, and per-cell style facets. */
export interface CellModel extends CellStyle {
  readonly row: number;
  readonly col: number;
  value: CellValue;
  note?: string | undefined;
}

/**
 * A serialisable snapshot of a worksheet's value and overlay content — its cells and their styles,
 * the column/row/page metadata, and the sheet-level overlays (merges, data validations, conditional
 * formattings, tables, the autofilter, protection). {@link Worksheet.model} exports one; assigning
 * it back reproduces that content. The getter and setter cover exactly the same fields, so a
 * `dst.model = src.model` round-trip drops none of it — an export field the import ignored would
 * silently lose data, the historical merge-loss failure this contract exists to prevent.
 *
 * Out of scope by design: content that carries workbook-level identity rather than pure sheet
 * state — anchored and background images (their bytes live on the {@link Workbook}), pivot tables
 * (their source references a live worksheet), and byte-preserved parts (charts, vector drawings,
 * slicers) kept verbatim for round-tripping. These stay with their source sheet; a model assignment
 * neither copies nor clears them.
 */
export interface WorksheetModel {
  state: WorksheetState['state'];
  tabColor: Color | undefined;
  properties: WorksheetProperties;
  outline: OutlineProperties;
  pageSetup: PageSetup;
  printOptions: PrintOptions;
  pageMargins: PageMargins;
  headerFooter: HeaderFooter;
  rowBreaks: PageBreak[];
  columnBreaks: PageBreak[];
  columns: {index: number; properties: ColumnProperties}[];
  rows: {number: number; properties: RowProperties}[];
  cells: CellModel[];
  merges: string[];
  dataValidations: DataValidationEntry[];
  conditionalFormattings: ConditionalFormatting[];
  tables: TableOptions[];
  autoFilter: AutoFilter | undefined;
  protection: SheetProtection | undefined;
}

// Sub-cell anchor geometry: a fractional grid coordinate resolves to the cell it floors to plus an
// EMU offset scaled by that cell's real size. Excel measures a column in characters of the default
// font (~7 px each at 96 DPI) and a row in points (1/72 inch); an unset size falls back to Excel's
// own defaults. Consumed by the class's `#columnWidthEmu`/`#rowHeightEmu` anchor-resolution methods.
const CHAR_WIDTH_PX = 7;
const EMU_PER_POINT = 12700;
const DEFAULT_COL_WIDTH_CHARS = 8.43;
const DEFAULT_ROW_HEIGHT_POINTS = 15;

export class Worksheet {
  readonly name: string;
  /** 1-based workbook-assigned id, stable for the sheet's lifetime. */
  readonly id: number;
  state: WorksheetState['state'];

  /**
   * Colour of the sheet's tab, as an ARGB/theme {@link Color}. `undefined` leaves the tab its
   * default colour; the writer emits no `<tabColor>` for an uncoloured sheet, so a round-trip
   * never fabricates one.
   */
  tabColor: Color | undefined;

  /** Sheet-level format defaults. Mutate in place: `sheet.properties.defaultRowHeight = 20`. */
  readonly properties: WorksheetProperties = {};

  /**
   * Outline summary-position flags. Mutate in place: `sheet.outline.summaryBelow = false`. Empty
   * means unset — the writer emits no `<outlinePr>` and a round-trip never fabricates one.
   */
  readonly outline: OutlineProperties = {};

  /**
   * The sheet's frozen-pane view. Empty (a normal view) emits no `<pane>`. Use {@link freeze} and
   * {@link unfreeze} for the common cases, or mutate in place for finer control.
   */
  readonly view: SheetView = {};

  /**
   * Print-scaling and orientation. Mutate in place: `sheet.pageSetup.fitToPage = true`. Empty means
   * unset — the writer emits neither `<pageSetUpPr>` nor `<pageSetup>` and a round-trip never
   * fabricates them.
   */
  readonly pageSetup: PageSetup = {};

  /**
   * Print-toggle flags (`<printOptions>`): centring, and whether headings/gridlines print. Mutate in
   * place: `sheet.printOptions.gridLines = true`. Empty means unset — the writer emits no element and
   * a round-trip never fabricates one.
   */
  readonly printOptions: PrintOptions = {};

  /** Print margins. Mutate in place: `sheet.pageMargins.left = 0.5`. Empty means unset. */
  readonly pageMargins: PageMargins = {};

  /** Page header/footer text. Mutate in place: `sheet.headerFooter.oddHeader = '&C&"..."'`. */
  readonly headerFooter: HeaderFooter = {};

  /**
   * Manual horizontal page breaks (`<rowBreaks>`): each break's `id` is a row the print layout splits
   * before. Mutate in place: `sheet.rowBreaks.push({id: 3})`. Empty means no manual row breaks and the
   * writer emits no `<rowBreaks>` element.
   */
  readonly rowBreaks: PageBreak[] = [];

  /**
   * Manual vertical page breaks (`<colBreaks>`): each break's `id` is a column the print layout splits
   * before. Mutate in place: `sheet.columnBreaks.push({id: 3})`. Empty means no manual column breaks and
   * the writer emits no `<colBreaks>` element.
   */
  readonly columnBreaks: PageBreak[] = [];

  // Row-major sparse storage: row index → (column index → cell). Keeping rows as the
  // outer key makes whole-row iteration cheap and mirrors how OOXML serializes
  // (`<row>` wrapping `<c>`).
  readonly #rows = new Map<number, Map<number, Cell>>();
  // Column and row metadata live apart from the grid so an empty-but-formatted line
  // (a hidden column, a tall header row with no data yet) costs no phantom cells.
  readonly #columns = new Map<number, ColumnProperties>();
  readonly #rowProperties = new Map<number, RowProperties>();
  // Tables, merged ranges, and anchored images are sheet-level overlays on the grid, not cell storage.
  readonly #tables: Table[] = [];
  readonly #pivotTables: PivotTable[] = [];
  // Pivot tables reconstructed from a loaded package (see io/xlsx/pivot-read.ts) — a read-only,
  // inspection-only view distinct from #pivotTables. A loaded pivot round-trips by byte-preservation
  // (#preservedReferences), which stays its sole emission authority; this collection is never emitted,
  // so exposing it cannot double-emit. Empty for a sheet authored from scratch.
  readonly #loadedPivotTables: ParsedPivotTable[] = [];
  readonly #merges: string[] = [];
  readonly #images: AnchoredImage[] = [];
  // A sheet background is a single workbook image tiled behind the grid — distinct from an anchored
  // drawing (it has no anchor and rides its own worksheet relationship, not a drawing part).
  #backgroundImageId: number | undefined;
  // Worksheet-level references to package content the model does not interpret (a vector-shape
  // drawing, a header/footer image), captured verbatim on read so a round-trip re-emits them rather
  // than dropping them. Empty for a sheet authored from scratch.
  readonly #preservedReferences: PreservedWorksheetReference[] = [];
  // Decoded rectangles parallel to #merges, kept so that addressing a covered cell can
  // resolve to its region's master without re-parsing the range string on every access, and
  // so that a new merge can be checked for overlap against the existing ones. Only fully-bounded
  // merges (a real cell block) get a rect; an unbounded whole-row/column merge is still declared
  // but participates in neither slave resolution nor overlap checking.
  readonly #mergeRects: MergeRect[] = [];
  // Data validations are a sheet-level overlay keyed by range, parallel to merges: the entries carry
  // the serialisable form, the rects the decoded ranges a cell lookup tests for containment.
  readonly #dataValidations: DataValidationEntry[] = [];
  readonly #dataValidationRects: {rects: readonly MergeRect[]; rule: DataValidation}[] = [];
  // Conditional formattings are a sheet-level overlay keyed by range, like data validations: each
  // block binds a set of rules to the area(s) it covers, layered by the rules' evaluation precedence.
  readonly #conditionalFormattings: ConditionalFormatting[] = [];
  // Sheet-level protection is a single overlay switch, absent until `protect` is called.
  #protection: SheetProtection | undefined;
  // The sheet's autofilter (range plus any per-column criteria), absent until one is set. A single
  // sheet-level overlay, distinct from a table's own autofilter; stored canonically so the
  // `<autoFilter>` element and the derived `_FilterDatabase` defined name always agree.
  #autoFilter: AutoFilter | undefined;

  // Structural-edit machinery (row/column splices), sharing this sheet's storage by reference. The
  // public spliceRows/spliceColumns/duplicateRow build the cells an insert introduces, then delegate
  // the shift arithmetic here.
  readonly #edits = new GridEdits({
    rows: this.#rows,
    rowProperties: this.#rowProperties,
    columns: this.#columns,
    merges: this.#merges,
    mergeRects: this.#mergeRects,
    tables: this.#tables,
    images: this.#images,
  });

  constructor(name: string, id: number, state: WorksheetState['state'] = 'visible') {
    this.name = name;
    this.id = id;
    this.state = state;
  }

  /**
   * Get the cell at an A1 reference, creating it on first access. The reference must
   * name both a column and a row (`"B3"`); a whole-row or whole-column reference is
   * not a cell and is rejected.
   *
   * Addressing a cell covered by a merged region resolves to that region's master
   * (top-left) cell, mirroring how a spreadsheet treats the merge as one cell: a value
   * or style written through a covered address lands on the master, and reading a
   * covered address returns the master's. Only the master ever holds an independent
   * value, so the serialized sheet stays well-formed (no stray value on a covered cell).
   *
   * @throws {SyntaxError} if the reference does not resolve to a single cell.
   */
  getCell(reference: string): Cell {
    const {col, row} = decodeAddress(reference);
    if (col === undefined || row === undefined) {
      throw new SyntaxError(
        `"${reference}" is not a single-cell reference — it omits a column or row`,
      );
    }
    const master = this.#masterOf(row, col);
    return this.#cellAt(master.row, master.col);
  }

  /** Whether a cell has been materialised at the given 1-based position. */
  hasCell(row: number, col: number): boolean {
    return this.#rows.get(row)?.has(col) ?? false;
  }

  /** The format properties for a 1-based row number if any were set, without materialising them —
   * the read-only peek {@link getRow} is not, so a serializer can render a row's attributes without
   * fabricating an empty record for every row it visits. */
  rowProperties(number: number): RowProperties | undefined {
    return this.#rowProperties.get(number);
  }

  /**
   * Drop a row's materialised cells and format properties, releasing its cell graph. The streaming
   * writer calls this the moment a row is serialised so peak memory stays bounded to the rows still
   * in flight rather than the whole sheet. Row *numbering* is the caller's concern: eviction lowers
   * {@link rowCount}, so an append-driven producer must track its own high-water mark rather than
   * lean on this sheet's used range.
   */
  evictRow(number: number): void {
    this.#rows.delete(number);
    this.#rowProperties.delete(number);
  }

  /**
   * Get the mutable format properties for a 1-based column index, creating the record
   * on first access. Setting properties here does not materialise any cells.
   *
   * @throws {RangeError} if the index is not a positive integer.
   */
  getColumn(index: number): ColumnProperties {
    if (!Number.isInteger(index) || index < 1) {
      throw new RangeError(`column ${index} is out of bounds — columns start at 1`);
    }
    let properties = this.#columns.get(index);
    if (properties === undefined) {
      properties = {};
      this.#columns.set(index, properties);
    }
    return properties;
  }

  /**
   * Get the mutable format properties for a 1-based row number, creating the record on
   * first access. This is row *metadata* (height, visibility, outline) — it does not
   * materialise any cells.
   *
   * @throws {RangeError} if the number is not a positive integer.
   */
  getRow(number: number): RowProperties {
    if (!Number.isInteger(number) || number < 1) {
      throw new RangeError(`row ${number} is out of bounds — rows start at 1`);
    }
    let properties = this.#rowProperties.get(number);
    if (properties === undefined) {
      properties = {};
      this.#rowProperties.set(number, properties);
    }
    return properties;
  }

  /**
   * The 1-based index of the last row carrying anything — data or its own formatting —
   * or 0 for an empty sheet. Spans gaps: a value in row 5 makes this 5 even if rows 2–4
   * are empty. This is the used-range extent, not a populated-row tally (see
   * {@link actualRowCount}).
   */
  get rowCount(): number {
    let last = 0;
    for (const [number, cols] of this.#rows) {
      if (number > last && this.#rowHasContent(cols)) last = number;
    }
    for (const number of this.#rowProperties.keys()) {
      if (number > last) last = number;
    }
    // A merged region occupies its whole rectangle even where the covered cells are empty, so a merge
    // extending past the last populated row still belongs to the used range.
    for (const rect of this.#mergeRects) {
      if (rect.bottom > last) last = rect.bottom;
    }
    return last;
  }

  /** The number of rows that hold at least one non-empty cell, ignoring gaps and formatting-only rows. */
  get actualRowCount(): number {
    let count = 0;
    for (const cols of this.#rows.values()) {
      if (this.#rowHasContent(cols)) count++;
    }
    return count;
  }

  // Whether any cell materialised in a row holds a value — the used-range test {@link rowCount} and
  // {@link actualRowCount} share. Short-circuits on the first non-empty cell rather than allocating the
  // row's values into a throwaway array to scan them.
  #rowHasContent(cols: Map<number, Cell>): boolean {
    for (const cell of cols.values()) {
      if (cell.value !== null) return true;
    }
    return false;
  }

  /**
   * The 1-based index of the last column carrying anything — a non-empty cell or its own format
   * properties — or 0 for an empty sheet. The used-range width, mirroring {@link rowCount} for the
   * other axis: a value in column E makes this 5 even if columns B–D are empty.
   */
  get columnCount(): number {
    let last = 0;
    for (const cols of this.#rows.values()) {
      for (const [col, cell] of cols) {
        if (cell.value !== null && col > last) last = col;
      }
    }
    for (const index of this.#columns.keys()) {
      if (index > last) last = index;
    }
    for (const rect of this.#mergeRects) {
      if (rect.right > last) last = rect.right;
    }
    return last;
  }

  /** The defined columns in ascending index order, each with its format properties. */
  *columns(): IterableIterator<{readonly index: number; readonly properties: ColumnProperties}> {
    for (const [index, properties] of [...this.#columns].sort(([a], [b]) => a - b)) {
      yield {index, properties};
    }
  }

  /**
   * The rows to serialise, in ascending row order: the union of rows holding cells and
   * rows holding only metadata (a hidden or grouped row need carry no data). Each yields
   * its materialised cells in ascending column order and its format properties, if any.
   * Mirrors how OOXML serialises (`<row>` wrapping `<c>`) and is the writer's row surface.
   */
  *rows(): IterableIterator<{
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
  }> {
    const numbers = new Set<number>([...this.#rows.keys(), ...this.#rowProperties.keys()]);
    for (const number of [...numbers].sort((a, b) => a - b)) {
      const cols = this.#rows.get(number);
      const cells = cols ? [...cols].sort(([a], [b]) => a - b).map(([, cell]) => cell) : [];
      yield {number, cells, properties: this.#rowProperties.get(number)};
    }
  }

  /**
   * Define a table over a range of this sheet. The table's shape invariants (a legal
   * name, at least one column, at least one row) are enforced here; conflicts with the
   * rest of the sheet (e.g. an overlapping merge) are the writer's concern.
   *
   * @throws {Error} if the name, columns, or geometry are invalid.
   */
  addTable(options: TableOptions): Table {
    const table = new Table(
      options,
      (row, col, value, style) => {
        const cell = this.#cellAt(row, col);
        cell.value = value;
        if (style !== undefined) applyCellStyle(cell, style);
      },
      // Insert one empty grid row at `row`; the splice re-pins this table (growing its data rows) and
      // shifts the totals row and everything below down by one.
      (row) => this.spliceRows(row, 0, []),
    );
    this.#tables.push(table);
    return table;
  }

  /** The tables defined on this sheet, in definition order. */
  get tables(): readonly Table[] {
    return this.#tables;
  }

  /** The table with the given name (case-sensitive, the identifier Excel uses), or `undefined`.
   * A table read back from a file is fully hydrated — its rows can be read and appended to. */
  getTable(name: string): Table | undefined {
    return this.#tables.find((table) => table.name === name);
  }

  /**
   * Add a pivot table to this (destination) sheet, summarising a source sheet's data. The source is
   * read once, now, so the pivot is a snapshot: later edits to the source do not change it. The
   * supported shape (one summed value field, at least one row and column field) is enforced here.
   *
   * @throws {Error} if the metric, fields, or source shape are unsupported.
   */
  addPivotTable(options: PivotTableOptions): PivotTable {
    const pivot = new PivotTable(options);
    this.#pivotTables.push(pivot);
    return pivot;
  }

  /** The pivot tables hosted on this sheet, in definition order. */
  get pivotTables(): readonly PivotTable[] {
    return this.#pivotTables;
  }

  /**
   * Register a pivot table reconstructed from a loaded package — the reader's counterpart to
   * {@link addPivotTable}. This records an inspectable, read-only view of a pivot the reader parsed
   * from its OOXML parts; the pivot itself round-trips by byte-preservation, so registering it here
   * only makes it visible via {@link loadedPivotTables} and never affects what the writer emits.
   */
  addLoadedPivotTable(pivot: ParsedPivotTable): void {
    this.#loadedPivotTables.push(pivot);
  }

  /**
   * Pivot tables reconstructed from a loaded package, in the order the reader found them — a
   * read-only inspection view (source range, field roles, value field, aggregation). A pivot
   * authored on this sheet via {@link addPivotTable} does not appear here; a pivot loaded from a
   * file does not appear in {@link pivotTables}. The loaded pivots re-emit verbatim through
   * byte-preservation, so this collection is never itself serialised.
   */
  get loadedPivotTables(): readonly ParsedPivotTable[] {
    return this.#loadedPivotTables;
  }

  /**
   * Anchor a workbook image (the id returned by {@link Workbook.addImage}) to this sheet. Two shapes:
   *
   * - **Two-cell**: `{tl, br}` spans the rectangle from the top-left grid point to the bottom-right,
   *   reflowing as the spanned cells resize. `editAs` (`oneCell` by default) tunes how it follows.
   * - **One-cell**: `{tl, ext}` pins the image at `tl` at a fixed pixel size that the grid never
   *   resizes. `ext` is in pixels and converts to EMUs internally.
   *
   * Grid points are 0-based (`{col: 0, row: 0}` is cell A1). A later row/column splice re-pins the
   * anchor to the same logical position.
   */
  addImage(
    imageId: number,
    anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs},
  ): void;
  addImage(
    imageId: number,
    anchor: {
      readonly tl: AnchorPoint;
      readonly ext: {readonly width: number; readonly height: number};
    },
  ): void;
  addImage(
    imageId: number,
    anchor:
      | {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs}
      | {readonly tl: AnchorPoint; readonly ext: {readonly width: number; readonly height: number}},
  ): void {
    if ('ext' in anchor) {
      const ext: Extent = {
        cx: Math.round(anchor.ext.width * PX_TO_EMU),
        cy: Math.round(anchor.ext.height * PX_TO_EMU),
      };
      this.#images.push({imageId, anchor: {from: this.#resolveAnchorPoint(anchor.tl), ext}});
      return;
    }
    const from = this.#resolveAnchorPoint(anchor.tl);
    const to = this.#resolveAnchorPoint(anchor.br);
    const twoCell: TwoCellAnchor =
      anchor.editAs !== undefined ? {from, to, editAs: anchor.editAs} : {from, to};
    this.#images.push({imageId, anchor: twoCell});
  }

  // Resolve a possibly-fractional anchor point to the cell it floors to plus a sub-cell EMU offset
  // scaled by that cell's real width/height, so `col: 3.5` lands halfway across column 3 regardless
  // of the column's size. An already-integer point keeps a zero offset (unless one was given).
  #resolveAnchorPoint(point: AnchorPoint): AnchorPoint {
    const col = Math.floor(point.col);
    const row = Math.floor(point.row);
    const colOff = (point.colOff ?? 0) + Math.round((point.col - col) * this.#columnWidthEmu(col));
    const rowOff = (point.rowOff ?? 0) + Math.round((point.row - row) * this.#rowHeightEmu(row));
    return {col, row, colOff, rowOff};
  }

  #columnWidthEmu(col: number): number {
    const width =
      this.#columns.get(col + 1)?.width ??
      this.properties.defaultColWidth ??
      DEFAULT_COL_WIDTH_CHARS;
    return Math.round(width * CHAR_WIDTH_PX * PX_TO_EMU);
  }

  #rowHeightEmu(row: number): number {
    const height =
      this.#rowProperties.get(row + 1)?.height ??
      this.properties.defaultRowHeight ??
      DEFAULT_ROW_HEIGHT_POINTS;
    return Math.round(height * EMU_PER_POINT);
  }

  /**
   * Anchor an image with a pre-built model anchor in the model's own units (EMUs). This is the
   * low-level primitive {@link addImage} builds on and the reader uses to re-pin an image parsed from
   * a drawing part without a lossy pixel round-trip.
   */
  addImageAnchor(imageId: number, anchor: ImageAnchor): void {
    this.#images.push({imageId, anchor});
  }

  /** Drop every anchor of the given workbook image from this sheet. The image stays registered on the
   * workbook — another sheet may still show it — so only this sheet's anchors are removed; the writer
   * then omits any media no sheet anchors any longer. */
  removeImage(imageId: number): void {
    const kept = this.#images.filter((image) => image.imageId !== imageId);
    replaceContents(this.#images, kept);
  }

  /** The images anchored to this sheet, in the order they were added. */
  get images(): readonly AnchoredImage[] {
    return this.#images;
  }

  /** Set this sheet's background image to a workbook image (the id {@link Workbook.addImage} returned).
   * The picture tiles behind the whole grid; it is not anchored to any cell. Passing a new id replaces
   * the previous background. */
  addBackgroundImage(imageId: number): void {
    this.#backgroundImageId = imageId;
  }

  /** Remove this sheet's background image, if any. The image stays registered on the workbook. */
  removeBackgroundImage(): void {
    this.#backgroundImageId = undefined;
  }

  /** The workbook image id set as this sheet's background, or `undefined` when it has none. */
  get backgroundImageId(): number | undefined {
    return this.#backgroundImageId;
  }

  /**
   * Record a worksheet-level reference to package content the model does not interpret, so the writer
   * re-emits it verbatim. Called by the reader when it meets a `<drawing>` holding only vector shapes
   * or a `<legacyDrawingHF>` header/footer image; not part of the authoring surface.
   */
  addPreservedReference(reference: PreservedWorksheetReference): void {
    this.#preservedReferences.push(reference);
  }

  /** The worksheet-level references to unmodeled package content preserved for round-tripping. */
  get preservedReferences(): readonly PreservedWorksheetReference[] {
    return this.#preservedReferences;
  }

  /**
   * Merge a range of cells (`"A1:B2"`). A range that overlaps an already-merged region is
   * rejected — Excel forbids overlapping merges and writes such geometry as a corrupt file.
   * Whole-row/column ranges (`"A:A"`) are unbounded, carry no rectangle, and are not overlap-checked.
   */
  mergeCells(range: string): void {
    const {top, left, bottom, right} = decodeRange(range);
    if (top !== undefined && left !== undefined && bottom !== undefined && right !== undefined) {
      const rect: MergeRect = {top, left, bottom, right};
      const clash = this.#mergeRects.find((existing) => rectsOverlap(existing, rect));
      if (clash) {
        throw new Error(`merged range "${range}" overlaps an existing merged region`);
      }
      this.#mergeRects.push(rect);
    }
    this.#merges.push(range);
  }

  /** The merged ranges on this sheet, in the order they were added. */
  get merges(): readonly string[] {
    return this.#merges;
  }

  /**
   * The sheet's autofilter — its range plus any per-column criteria — or `undefined` when the sheet
   * carries none. Setting one turns on the header-row filter dropdowns Excel draws over the range;
   * the writer emits both the sheet's `<autoFilter>` element and the hidden `_FilterDatabase` defined
   * name Excel derives from it. Setting `undefined` clears the filter.
   *
   * A bare range string is the ergonomic common case — `sheet.autoFilter = 'A1:C10'` for dropdowns
   * with no active criteria; pass an {@link AutoFilter} object to narrow columns. Either way the
   * value is normalised on assignment (range to canonical `A1:C10` form) and the getter returns the
   * structured object. The range must be a bounded rectangle — a whole-row/column reference is not a
   * filterable region and is rejected.
   */
  get autoFilter(): AutoFilter | undefined {
    return this.#autoFilter;
  }

  set autoFilter(filter: string | AutoFilter | undefined) {
    this.#autoFilter = filter === undefined ? undefined : canonicalizeAutoFilter(filter);
  }

  /**
   * Remove a merged range previously added with {@link mergeCells}, returning whether a merge with
   * that exact range string existed. The covering rectangle is dropped alongside it, so a cell the
   * merge had masked addresses independently again. The inverse of {@link mergeCells}.
   */
  unmergeCells(range: string): boolean {
    const index = this.#merges.indexOf(range);
    if (index === -1) return false;
    this.#merges.splice(index, 1);
    const {top, left, bottom, right} = decodeRange(range);
    if (top !== undefined && left !== undefined && bottom !== undefined && right !== undefined) {
      const rectIndex = this.#mergeRects.findIndex(
        (r) => r.top === top && r.left === left && r.bottom === bottom && r.right === right,
      );
      if (rectIndex !== -1) this.#mergeRects.splice(rectIndex, 1);
    }
    return true;
  }

  /**
   * Attach a data validation to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a
   * space-separated `sqref` of several ranges). The rule is stored once against the range, not copied
   * per covered cell, so a whole-column dropdown stays a single entry. A cell inside the range reports
   * the rule through {@link dataValidationAt}.
   *
   * Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form
   * (`<x14:dataValidation>`) — the carrier Excel uses for a list source on another sheet and other
   * shapes the standard element cannot express. The reader sets it for a rule found in that form so a
   * round-trip writes it back there instead of silently corrupting the cross-sheet reference.
   */
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    // One defensive copy, shared by the serialisable entry and the lookup index, so the getter never
    // hands back a reference into the caller's object.
    const stored = cloneDataValidation(rule);
    const entry: DataValidationEntry = {sqref, rule: stored};
    if (options.extended) entry.extended = true;
    this.#dataValidations.push(entry);
    this.#dataValidationRects.push({rects: decodeSqrefRects(sqref), rule: stored});
  }

  /** The data validations on this sheet, each bound to its target range, in insertion order. */
  get dataValidations(): readonly DataValidationEntry[] {
    return this.#dataValidations;
  }

  /**
   * Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref` — one
   * range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one
   * rule set. The block is stored once against the range, defensively copied so the getter never hands
   * back a reference into the caller's object.
   */
  addConditionalFormatting(formatting: ConditionalFormatting): void {
    this.#conditionalFormattings.push(cloneConditionalFormatting(formatting));
  }

  /** The conditional formattings on this sheet, each bound to its target range, in insertion order. */
  get conditionalFormattings(): readonly ConditionalFormatting[] {
    return this.#conditionalFormattings;
  }

  /**
   * The validation covering a cell, or `undefined` when none does. The first added rule whose range
   * contains the cell wins, mirroring how a spreadsheet resolves overlapping validations.
   */
  dataValidationAt(reference: string): DataValidation | undefined {
    const {col, row} = decodeAddress(reference);
    if (col === undefined || row === undefined) return undefined;
    for (const {rects, rule} of this.#dataValidationRects) {
      for (const rect of rects) {
        if (col >= rect.left && col <= rect.right && row >= rect.top && row <= rect.bottom) {
          return rule;
        }
      }
    }
    return undefined;
  }

  /**
   * Remove `count` rows starting at the 1-based `start`, then insert the given rows in their place.
   * Rows below the edit shift by `inserts.length - count`: a delete pulls the tail up, an insert
   * pushes it down, and doing both at once is a replace. Each inserted row is an array of cell
   * values, one per column from A. A `count` larger than the rows present simply clears the tail —
   * it never silently becomes a no-op. Cells carry their full style to the shifted position, and
   * merged ranges shift with the rows they cover.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative.
   */
  spliceRows(start: number, count: number, ...inserts: CellValue[][]): void {
    if (!Number.isInteger(start) || start < 1) {
      throw new RangeError(`splice start ${start} is out of bounds — rows start at 1`);
    }
    if (!Number.isInteger(count) || count < 0) {
      throw new RangeError(`splice count ${count} is invalid — it must be a non-negative integer`);
    }
    const inserted = inserts.map((values, i) => {
      const row = new Map<number, Cell>();
      values.forEach((value, index) => {
        const cell = new Cell(start + i, index + 1);
        cell.value = value;
        row.set(index + 1, cell);
      });
      return row;
    });
    this.#edits.spliceRows(start, count, inserted);
  }

  /**
   * Insert one row of `values` at the 1-based `pos`, shifting the rows at and below it down by one.
   * Shorthand for {@link spliceRows}`(pos, 0, values)`.
   *
   * @throws {RangeError} if `pos` is not a positive integer.
   */
  insertRow(pos: number, values: CellValue[]): void {
    this.spliceRows(pos, 0, values);
  }

  /**
   * Append a row of `values` after the last used row, returning the cells it materialised.
   * The append point is {@link rowCount}` + 1`, so the row lands below every row that holds
   * data or its own formatting — never overwriting existing content, unlike {@link insertRow},
   * which shifts and needs a position. Unlike {@link spliceRows}, appending shifts nothing, so
   * it never disturbs merges or the rows above.
   *
   * A row takes either shape: a positional array whose values map to columns from A — a hole in a
   * sparse array (`['a', , 'c']`) leaves that column untouched — or a keyed object whose values
   * land under the columns carrying the matching {@link ColumnProperties.key}.
   */
  addRow(values: RowInput): Cell[] {
    return this.addRows([values])[0] ?? [];
  }

  /**
   * Append several rows after the last used row in one call, returning the cells materialised
   * for each. The rows stack in order — the first lands at {@link rowCount}` + 1`, the next
   * directly below it — so a later row never collides with an earlier one even when both are
   * value-less. Each row is an array or a keyed object independently, so a mixed batch is fine.
   * The bulk form of {@link addRow}.
   */
  addRows(rows: RowInput[]): Cell[][] {
    let number = this.rowCount;
    return rows.map((values) => {
      number += 1;
      const cells: Cell[] = [];
      const place = (col: number, value: CellValue): void => {
        const cell = this.#cellAt(number, col);
        cell.value = value;
        cells.push(cell);
      };
      // Array.isArray, not `instanceof Array`: a row built in another realm (a vm context, a browser
      // iframe) is still an array but fails the identity check, and would then be walked as a keyed
      // object — placing nothing.
      if (Array.isArray(values)) {
        values.forEach((value, index) => {
          if (value !== undefined) place(index + 1, value);
        });
      } else {
        for (const [key, value] of Object.entries(values))
          place(this.#columnIndexByKey(key), value);
      }
      return cells;
    });
  }

  /**
   * Freeze the top `ySplit` rows and left `xSplit` columns in place; the rest of the sheet scrolls
   * beneath them. `freeze(1)` pins a header row; `freeze(0, 1)` pins the first column. Passing both
   * zero clears the freeze (equivalent to {@link unfreeze}).
   *
   * @throws {RangeError} if either split is a negative or non-integer count.
   */
  freeze(ySplit = 1, xSplit = 0): void {
    if (!Number.isInteger(ySplit) || ySplit < 0 || !Number.isInteger(xSplit) || xSplit < 0) {
      throw new RangeError(
        `freeze splits must be non-negative integers; got ySplit=${ySplit}, xSplit=${xSplit}`,
      );
    }
    if (ySplit === 0 && xSplit === 0) {
      this.unfreeze();
      return;
    }
    this.view.state = 'frozen';
    this.view.xSplit = xSplit;
    this.view.ySplit = ySplit;
    this.view.topLeftCell = encodeAddress(xSplit + 1, ySplit + 1);
  }

  /** Clear any frozen split, returning the sheet to a normal (fully scrolling) view. */
  unfreeze(): void {
    this.view.state = 'normal';
    delete this.view.xSplit;
    delete this.view.ySplit;
    delete this.view.topLeftCell;
  }

  /** The 1-based index of the column carrying `key` (see {@link ColumnProperties.key}). */
  #columnIndexByKey(key: string): number {
    for (const [index, properties] of this.#columns) {
      if (properties.key === key) return index;
    }
    throw new Error(`no column is keyed ${JSON.stringify(key)} — set getColumn(n).key first`);
  }

  /**
   * Copy the row at the 1-based `start`, `count` times. With `insert` (the default) the copies are
   * inserted directly after the source, shifting the rows below — and any merged range there — down
   * by `count`; otherwise the copies overwrite the rows immediately below without shifting. Each
   * copy is a faithful duplicate of the source's values and per-cell styles, and carries no merge of
   * its own, so a range can be merged onto a duplicated row afterwards.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative.
   */
  duplicateRow(start: number, count = 1, insert = true): void {
    if (!Number.isInteger(start) || start < 1) {
      throw new RangeError(`duplicate start ${start} is out of bounds — rows start at 1`);
    }
    if (!Number.isInteger(count) || count < 0) {
      throw new RangeError(
        `duplicate count ${count} is invalid — it must be a non-negative integer`,
      );
    }
    const source = this.#rows.get(start);
    const snapshot = (destRow: number): Map<number, Cell> => {
      const row = new Map<number, Cell>();
      if (source) {
        for (const [col, cell] of source) {
          const copy = new Cell(destRow, col);
          copyCellContent(cell, copy);
          row.set(col, copy);
        }
      }
      return row;
    };
    if (insert) {
      const copies = Array.from({length: count}, () => snapshot(start));
      this.#edits.spliceRows(start + 1, 0, copies);
    } else {
      for (let i = 1; i <= count; i++) this.#rows.set(start + i, snapshot(start + i));
    }
  }

  /**
   * Remove `count` columns starting at the 1-based `start`, then insert the given columns in their
   * place — the column analog of {@link spliceRows}. Columns to the right shift by
   * `inserts.length - count`, keeping their values and styles, and a merged range lying wholly to
   * the right of the edit re-anchors to its new columns. Each inserted column is an array of values
   * indexed by row (index 0 → row 1); an empty array inserts a blank column.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative.
   */
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void {
    if (!Number.isInteger(start) || start < 1) {
      throw new RangeError(`splice start ${start} is out of bounds — columns start at 1`);
    }
    if (!Number.isInteger(count) || count < 0) {
      throw new RangeError(`splice count ${count} is invalid — it must be a non-negative integer`);
    }
    this.#edits.spliceColumns(start, count, inserts);
  }

  /**
   * A snapshot of this sheet's value and overlay content (see {@link WorksheetModel}). Reading it and
   * assigning it onto another sheet — `dst.model = src.model` — reproduces the source: merges, cells
   * and their styles, column/row metadata, tables, the autofilter, protection, and the page setup all
   * survive, because the getter emits and the setter consumes exactly the same fields. Identity
   * (`name`, `id`) is not part of the model and is never touched by assignment; nor are attached parts
   * that carry workbook-level identity (images, pivots, byte-preserved charts/drawings) — see
   * {@link WorksheetModel} for that boundary.
   */
  get model(): WorksheetModel {
    const cells: CellModel[] = [];
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) cells.push(cellToModel(cell));
    }
    return {
      state: this.state,
      tabColor: this.tabColor,
      properties: {...this.properties},
      outline: {...this.outline},
      pageSetup: {...this.pageSetup},
      printOptions: {...this.printOptions},
      pageMargins: {...this.pageMargins},
      headerFooter: {...this.headerFooter},
      rowBreaks: this.rowBreaks.map((brk) => ({...brk})),
      columnBreaks: this.columnBreaks.map((brk) => ({...brk})),
      columns: [...this.#columns].map(([index, properties]) => ({
        index,
        properties: {...properties},
      })),
      rows: [...this.#rowProperties].map(([number, properties]) => ({
        number,
        properties: {...properties},
      })),
      cells,
      merges: [...this.#merges],
      dataValidations: this.#dataValidations.map(({sqref, rule, extended}) => ({
        sqref,
        rule: cloneDataValidation(rule),
        ...(extended ? {extended: true} : {}),
      })),
      conditionalFormattings: this.#conditionalFormattings.map(cloneConditionalFormatting),
      tables: this.#tables.map((table) => table.options),
      autoFilter: this.#autoFilter,
      protection: this.#protection,
    };
  }

  // Empty every collection the model round-trips, so a subsequent replay leaves no residue from
  // whatever the sheet held before. Images, pivots, and byte-preserved parts carry workbook-level
  // identity, are excluded from the model, and so are deliberately left untouched here.
  #resetContent(): void {
    this.#rows.clear();
    this.#columns.clear();
    this.#rowProperties.clear();
    this.#merges.length = 0;
    this.#mergeRects.length = 0;
    this.#dataValidations.length = 0;
    this.#dataValidationRects.length = 0;
    this.#conditionalFormattings.length = 0;
    this.#tables.length = 0;
  }

  // Assigning a model replaces this sheet's content wholesale — the sheet becomes the model, with no
  // residue from whatever it held before. Cells are placed at their exact positions (bypassing merge
  // resolution) and merges re-applied after, so a slave's value cannot be misrouted during the load.
  set model(model: WorksheetModel) {
    this.#resetContent();
    this.state = model.state;
    this.tabColor = model.tabColor;
    overwrite(this.properties, model.properties);
    overwrite(this.outline, model.outline);
    overwrite(this.pageSetup, model.pageSetup);
    overwrite(this.printOptions, model.printOptions);
    overwrite(this.pageMargins, model.pageMargins);
    overwrite(this.headerFooter, model.headerFooter);
    replaceContents(
      this.rowBreaks,
      model.rowBreaks.map((brk) => ({...brk})),
    );
    replaceContents(
      this.columnBreaks,
      model.columnBreaks.map((brk) => ({...brk})),
    );

    this.#protection = model.protection;

    for (const {index, properties} of model.columns)
      Object.assign(this.getColumn(index), properties);
    for (const {number, properties} of model.rows) Object.assign(this.getRow(number), properties);
    for (const cellModel of model.cells) {
      copyCellContent(cellModel, this.#cellAt(cellModel.row, cellModel.col));
    }
    for (const range of model.merges) this.mergeCells(range);
    for (const {sqref, rule, extended} of model.dataValidations) {
      this.addDataValidation(sqref, rule, extended ? {extended: true} : {});
    }
    for (const formatting of model.conditionalFormattings)
      this.addConditionalFormatting(formatting);
    for (const options of model.tables) this.addTable(options);
    // Assigning through the setter (rather than the private field) re-canonicalises and, on undefined,
    // clears any autofilter the destination held — the wholesale-replace contract, no residue.
    this.autoFilter = model.autoFilter;
  }

  /**
   * Protect the sheet, making the per-cell `locked`/`hidden` flags enforceable. Without a
   * password the protection is a soft lock any consumer can lift; with one, the password is
   * salted and hashed on the spot (the plaintext is never retained) so lifting the protection
   * requires re-supplying it. `options` names which operations stay available to a user while
   * the sheet is protected; anything unspecified falls to Excel's default for that operation.
   *
   * Re-protecting replaces any prior protection; {@link unprotect} clears it.
   */
  protect(password?: string, options: SheetProtectionOptions = {}): void {
    const {spinCount, ...flags} = options;
    const protection: {flags: SheetProtection['flags']; credential?: SheetProtectionCredential} = {
      flags,
    };
    if (password !== undefined && password !== '') {
      protection.credential = deriveCredential(password, spinCount);
    }
    this.#protection = protection;
  }

  /** Remove any protection previously set by {@link protect}. */
  unprotect(): void {
    this.#protection = undefined;
  }

  /**
   * Reinstate an already-derived protection state — the deserialization counterpart to
   * {@link protect}. A loaded `<sheetProtection>` carries its credential in finished agile form
   * (algorithm, hash, salt, spin count) with no recoverable plaintext password, so the reader
   * restores that credential verbatim rather than re-hashing. Use {@link protect} to protect from
   * a plaintext password; use this only to carry a parsed protection back into the model.
   */
  restoreProtection(protection: SheetProtection): void {
    this.#protection = protection;
  }

  /** The sheet's protection, or `undefined` if the sheet is unprotected. */
  get protection(): SheetProtection | undefined {
    return this.#protection;
  }

  // Resolve a position to the master (top-left) of the merged region covering it, or to
  // itself when no region does. First covering region wins; overlaps are rejected in
  // `mergeCells`, so at most one region ever applies.
  #masterOf(row: number, col: number): {row: number; col: number} {
    for (const rect of this.#mergeRects) {
      if (row >= rect.top && row <= rect.bottom && col >= rect.left && col <= rect.right) {
        return {row: rect.top, col: rect.left};
      }
    }
    return {row, col};
  }

  #cellAt(row: number, col: number): Cell {
    let cols = this.#rows.get(row);
    if (cols === undefined) {
      cols = new Map<number, Cell>();
      this.#rows.set(row, cols);
    }
    let cell = cols.get(col);
    if (cell === undefined) {
      cell = new Cell(row, col);
      cols.set(col, cell);
    }
    return cell;
  }
}
