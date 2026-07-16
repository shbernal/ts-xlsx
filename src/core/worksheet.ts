// A worksheet: a sparsely-populated grid of cells addressed by A1 reference.
//
// Storage is sparse by construction — a spreadsheet is mostly empty, so cells
// materialise on first access and only occupied positions cost memory. Column and
// row metadata (widths, heights, visibility, outline grouping) are stored apart from
// the cell grid, because a column or row can carry formatting while holding no cells.
// Merges and views layer on in later slices.

import {decodeAddress, decodeRange, encodeAddress} from './address.ts';
import {Cell} from './cell.ts';
import {
  cloneConditionalFormatting,
  type ConditionalFormatting,
} from './conditional-formatting.ts';
import {
  cloneDataValidation,
  type DataValidation,
  type DataValidationEntry,
} from './data-validation.ts';
import type {AnchoredImage, AnchorPoint} from './image.ts';
import {
  deriveCredential,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionOptions,
} from './protection.ts';
import type {Alignment, Border, Color, Fill, Font, Protection} from './style.ts';
import {Table, type TableOptions} from './table.ts';
import {type CellValue, isSharedFormulaValue, type SharedFormulaValue} from './value.ts';

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
 * Print-scaling and orientation settings. These map onto two OOXML elements: `fitToPage` is the
 * `<pageSetUpPr>` flag (a `<sheetPr>` child) that switches Excel from fixed-zoom to fit-to-page
 * scaling, while the rest are `<pageSetup>` attributes. Excel honours `scale` only when `fitToPage`
 * is off and the `fitToWidth`/`fitToHeight` page counts only when it is on, but the model carries
 * whatever the author set — an unset field is omitted so a round-trip never fabricates one. An
 * empty object emits neither element.
 */
export interface PageSetup {
  /** Switch to fit-to-page scaling. Emitted as `<pageSetUpPr fitToPage="1">`. */
  fitToPage?: boolean;
  /** Pages wide to fit onto; `0` means "unbounded" (fit only by height). */
  fitToWidth?: number;
  /** Pages tall to fit onto; `0` means "unbounded" (fit only by width). */
  fitToHeight?: number;
  /** Fixed print zoom as a percentage; Excel honours it only when `fitToPage` is off. */
  scale?: number;
  /** Paper orientation. */
  orientation?: 'portrait' | 'landscape';
  /** Order pages are numbered/printed in across a multi-page sheet. */
  pageOrder?: 'downThenOver' | 'overThenDown';
  /**
   * Paper size as Excel's 1-based enumeration index (e.g. `9` = A4, `1` = US Letter). Carried as an
   * opaque integer — the model does not map it to physical dimensions, only preserves whatever the
   * author or source file set.
   */
  paperSize?: number;
  /**
   * The printer-settings blob a source file bound to this sheet's `<pageSetup>` via an `r:id`
   * relationship, held verbatim. Excel stores the platform-specific `DEVMODE` (paper tray, duplex,
   * DPI, …) in this opaque binary part; the model does not interpret it, only round-trips the exact
   * bytes so re-writing a file that carried one does not silently drop the user's print configuration.
   */
  printerSettings?: Uint8Array;
}

/**
 * Print margins, in inches. OOXML's `<pageMargins>` requires all six to be present, but
 * the model stores only what the caller set; the writer fills the untouched ones with
 * valid defaults. An empty object means the element is omitted entirely.
 */
export interface PageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

/**
 * Page header/footer text, one string per page class. Excel only honours the even- and
 * first-page variants when the writer also sets the gating flags (`differentOddEven`,
 * `differentFirst`); the writer derives those from which variants are present. An empty
 * object means the element is omitted entirely.
 */
export interface HeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}

/**
 * Per-column formatting. A column may exist purely to carry these, with no cells. The style
 * facets are *defaults* for the column's cells: a cell that sets a facet of its own wins, but
 * one that leaves a facet unset inherits the column's — the same precedence Excel applies, and
 * symmetric with how a {@link RowProperties} fill defaults a row's cells.
 */
export interface ColumnProperties {
  /** Column width in character units. */
  width?: number;
  /** Whether the column is hidden. */
  hidden?: boolean;
  /** Number-format code applied to the column's cells that carry no format of their own. */
  numFmt?: string;
  /** Background fill applied to the column's cells that carry no fill of their own. */
  fill?: Fill;
  /** Font applied to the column's cells that carry no font of their own. */
  font?: Partial<Font>;
  /** Border applied to the column's cells that carry no border of their own. */
  border?: Border;
  /** Alignment applied to the column's cells that carry no alignment of their own. */
  alignment?: Alignment;
  /** Protection applied to the column's cells that carry no protection of their own. */
  protection?: Protection;
  /** Outline (grouping) depth; 0 or absent means ungrouped. */
  outlineLevel?: number;
  /** Whether this column is the collapsed summary of an outline group. */
  collapsed?: boolean;
}

/** A merged region as inclusive 1-based grid bounds. */
interface MergeRect {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** Whether two inclusive grid rectangles share at least one cell. */
function rectsOverlap(a: MergeRect, b: MergeRect): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** Decode an OOXML `sqref` (one or more space-separated ranges) into containment rectangles. A whole
 * column or row leaves one axis unbounded, so its missing edges open to `Infinity` rather than
 * clamping — a cell anywhere down the column still resolves inside it. */
function decodeSqrefRects(sqref: string): MergeRect[] {
  const rects: MergeRect[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (part === '') continue;
    const {top, left, bottom, right} = decodeRange(part);
    rects.push({
      top: top ?? 1,
      left: left ?? 1,
      bottom: bottom ?? Infinity,
      right: right ?? Infinity,
    });
  }
  return rects;
}

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
export interface CellModel {
  readonly row: number;
  readonly col: number;
  value: CellValue;
  fill?: Fill | undefined;
  numFmt?: string | undefined;
  font?: Partial<Font> | undefined;
  border?: Border | undefined;
  alignment?: Alignment | undefined;
  protection?: Protection | undefined;
  note?: string | undefined;
}

/**
 * A serialisable snapshot of a worksheet's transferable content — everything that defines the
 * sheet apart from its identity (`name`, `id`). {@link Worksheet.model} exports one; assigning it
 * back reproduces the sheet. The getter and setter cover exactly the same fields, so a
 * `dst.model = src.model` round-trip drops nothing — an export field the import ignored would
 * silently lose data, the historical merge-loss failure this contract exists to prevent.
 */
export interface WorksheetModel {
  state: WorksheetState['state'];
  tabColor: Color | undefined;
  properties: WorksheetProperties;
  outline: OutlineProperties;
  pageSetup: PageSetup;
  pageMargins: PageMargins;
  headerFooter: HeaderFooter;
  columns: {index: number; properties: ColumnProperties}[];
  rows: {number: number; properties: RowProperties}[];
  cells: CellModel[];
  merges: string[];
  dataValidations: DataValidationEntry[];
  conditionalFormattings: ConditionalFormatting[];
  tables: TableOptions[];
  protection: SheetProtection | undefined;
}

// Replace a mutable container's contents in place: a worksheet's `properties`/`pageSetup`/
// `pageMargins`/`headerFooter` are readonly fields holding mutable objects, so importing a model must overwrite
// them rather than reassign — and clear any stale key the incoming model does not carry.
function overwrite<T extends object>(target: T, source: T): void {
  const bag = target as unknown as Record<string, unknown>;
  for (const key of Object.keys(bag)) delete bag[key];
  Object.assign(target, source);
}

// Copy a cell's value and every style facet onto another cell. Used when a structural edit shifts a
// cell to a new position: `Cell` fixes its `(row, col)` at construction, so a shifted cell is a fresh
// cell at the new coordinates carrying the original's content. Facet objects are passed by reference —
// safe under the copy-on-write style model (facet setters replace, never mutate), so the source and
// its shifted copy never alias each other's style through a shared object.
function copyCellContent(source: Cell, target: Cell): void {
  target.value = source.value;
  target.fill = source.fill;
  target.numFmt = source.numFmt;
  target.font = source.font;
  target.border = source.border;
  target.alignment = source.alignment;
  target.protection = source.protection;
  target.note = source.note;
}

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
   * Print-scaling and orientation. Mutate in place: `sheet.pageSetup.fitToPage = true`. Empty means
   * unset — the writer emits neither `<pageSetUpPr>` nor `<pageSetup>` and a round-trip never
   * fabricates them.
   */
  readonly pageSetup: PageSetup = {};

  /** Print margins. Mutate in place: `sheet.pageMargins.left = 0.5`. Empty means unset. */
  readonly pageMargins: PageMargins = {};

  /** Page header/footer text. Mutate in place: `sheet.headerFooter.oddHeader = '&C&"..."'`. */
  readonly headerFooter: HeaderFooter = {};

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
  readonly #merges: string[] = [];
  readonly #images: AnchoredImage[] = [];
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
      throw new SyntaxError(`"${reference}" is not a single-cell reference — it omits a column or row`);
    }
    const master = this.#masterOf(row, col);
    return this.#cellAt(master.row, master.col);
  }

  /** Whether a cell has been materialised at the given 1-based position. */
  hasCell(row: number, col: number): boolean {
    return this.#rows.get(row)?.has(col) ?? false;
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
      if (number > last && [...cols.values()].some(cell => cell.value !== null)) last = number;
    }
    for (const number of this.#rowProperties.keys()) {
      if (number > last) last = number;
    }
    return last;
  }

  /** The number of rows that hold at least one non-empty cell, ignoring gaps and formatting-only rows. */
  get actualRowCount(): number {
    let count = 0;
    for (const cols of this.#rows.values()) {
      if ([...cols.values()].some(cell => cell.value !== null)) count++;
    }
    return count;
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
    const table = new Table(options);
    this.#tables.push(table);
    return table;
  }

  /** The tables defined on this sheet, in definition order. */
  get tables(): readonly Table[] {
    return this.#tables;
  }

  /**
   * Anchor a workbook image (the id returned by {@link Workbook.addImage}) to this sheet, spanning
   * the rectangle from the top-left grid point `tl` to the bottom-right `br`. Both points are
   * 0-based (`{col: 0, row: 0}` is cell A1). The anchor moves and reflows with the cells it spans, so
   * a later row/column splice re-pins it to the same logical position.
   */
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void {
    this.#images.push({imageId, anchor: {from: anchor.tl, to: anchor.br}});
  }

  /** The images anchored to this sheet, in the order they were added. */
  get images(): readonly AnchoredImage[] {
    return this.#images;
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
      const clash = this.#mergeRects.find(existing => rectsOverlap(existing, rect));
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
    this.#spliceRowRange(start, count, inserted);
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
   * which shifts and needs a position. Values map to columns from A; a hole in a sparse array
   * (`['a', , 'c']`) leaves that column untouched. Unlike {@link spliceRows}, appending shifts
   * nothing, so it never disturbs merges or the rows above.
   */
  addRow(values: CellValue[]): Cell[] {
    return this.addRows([values])[0] ?? [];
  }

  /**
   * Append several rows after the last used row in one call, returning the cells materialised
   * for each. The rows stack in order — the first lands at {@link rowCount}` + 1`, the next
   * directly below it — so a later row never collides with an earlier one even when both are
   * value-less. The bulk form of {@link addRow}.
   */
  addRows(rows: CellValue[][]): Cell[][] {
    let number = this.rowCount;
    return rows.map(values => {
      number += 1;
      const cells: Cell[] = [];
      values.forEach((value, index) => {
        const cell = this.#cellAt(number, index + 1);
        cell.value = value;
        cells.push(cell);
      });
      return cells;
    });
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
      throw new RangeError(`duplicate count ${count} is invalid — it must be a non-negative integer`);
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
      this.#spliceRowRange(start + 1, 0, copies);
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
    const delta = inserts.length - count;
    for (const [row, cols] of this.#rows) {
      const shifted = new Map<number, Cell>();
      for (const [col, cell] of cols) {
        if (col < start) {
          shifted.set(col, cell);
        } else if (col >= start + count) {
          const dest = col + delta;
          const moved = new Cell(row, dest);
          copyCellContent(cell, moved);
          shifted.set(dest, moved);
        }
      }
      inserts.forEach((values, i) => {
        const value = values[row - 1];
        if (value !== undefined) {
          const cell = new Cell(row, start + i);
          cell.value = value;
          shifted.set(start + i, cell);
        }
      });
      this.#rows.set(row, shifted);
    }
    this.#shiftLineProperties(this.#columns, start, count, delta);
    this.#shiftMerges(start, count, inserts.length, 'col');
    this.#shiftTables('col', start, count, delta);
    this.#shiftImages('col', start, count, delta);
    this.#reanchorSharedFormulas('col', start, count, delta);
  }

  // Apply a delete-then-insert to the row grid: surviving rows below the edit shift by
  // `inserts.length - count`, deleted rows drop out, and the pre-built inserted rows land at `start`.
  // Row metadata and merged ranges shift the same way, so a formatting-only row or a covered merge
  // stays aligned with the data it describes.
  #spliceRowRange(start: number, count: number, inserted: Map<number, Cell>[]): void {
    const delta = inserted.length - count;
    const shifted = new Map<number, Map<number, Cell>>();
    for (const [row, cols] of this.#rows) {
      if (row < start) shifted.set(row, cols);
      else if (row >= start + count) shifted.set(row + delta, this.#relocateRow(cols, row + delta));
    }
    inserted.forEach((cols, i) => shifted.set(start + i, this.#relocateRow(cols, start + i)));
    this.#rows.clear();
    for (const [row, cols] of shifted) this.#rows.set(row, cols);

    this.#shiftLineProperties(this.#rowProperties, start, count, delta);
    this.#shiftMerges(start, count, inserted.length, 'row');
    this.#shiftTables('row', start, count, delta);
    this.#shiftImages('row', start, count, delta);
    this.#reanchorSharedFormulas('row', start, count, delta);
  }

  // Rebuild a row's cells at a new row index. `Cell` fixes its position at construction, so a moved
  // row is a fresh set of cells at `destRow` carrying the originals' content.
  #relocateRow(cols: Map<number, Cell>, destRow: number): Map<number, Cell> {
    const moved = new Map<number, Cell>();
    for (const [col, cell] of cols) {
      if (cell.row === destRow) {
        moved.set(col, cell);
      } else {
        const copy = new Cell(destRow, col);
        copyCellContent(cell, copy);
        moved.set(col, copy);
      }
    }
    return moved;
  }

  // Re-anchor shared-formula clones through a splice on the given axis. A clone stores its master's
  // absolute address; when the splice shifts the master, that stored address goes stale and the writer
  // would reject the clone as orphaned. Applying the same shift the grid used keeps each clone pointed
  // at its master's new cell. A master whose axis coordinate falls in the deleted span clamps to the
  // cut line like a merge edge — a genuinely orphaned clone the writer then reports legibly.
  #reanchorSharedFormulas(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const shift = (v: number): number => (v < start ? v : v >= start + count ? v + delta : start);
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        if (!isSharedFormulaValue(value)) continue;
        const master = decodeAddress(value.sharedFormula);
        if (master.col === undefined || master.row === undefined) continue;
        const anchored =
          axis === 'row'
            ? encodeAddress(master.col, shift(master.row))
            : encodeAddress(shift(master.col), master.row);
        if (anchored === value.sharedFormula) continue;
        const reanchored: SharedFormulaValue = {...value, sharedFormula: anchored};
        cell.value = reanchored;
      }
    }
  }

  // Shift a line-metadata map (row properties keyed by row, or column properties keyed by column)
  // through a splice: entries before the edit stay, entries within the deleted span drop, entries
  // after shift by `delta`. Mutates the map in place.
  #shiftLineProperties<T>(map: Map<number, T>, start: number, count: number, delta: number): void {
    const shifted = new Map<number, T>();
    for (const [index, value] of map) {
      if (index < start) shifted.set(index, value);
      else if (index >= start + count) shifted.set(index + delta, value);
    }
    map.clear();
    for (const [index, value] of shifted) map.set(index, value);
  }

  // Re-anchor merged ranges through a row or column splice. A range wholly before the edit is
  // untouched; one wholly after shifts by `nInserts - count`; one whose covered rows/columns are
  // entirely deleted is dropped. A range straddling the cut is a genuinely ambiguous geometry — its
  // edges are clamped to the cut line as a best effort. Unbounded whole-row/column merges carry no
  // rectangle and pass through unchanged.
  #shiftMerges(start: number, count: number, nInserts: number, axis: 'row' | 'col'): void {
    const delta = nInserts - count;
    const shift = (v: number): number => (v < start ? v : v >= start + count ? v + delta : start);
    const merges: string[] = [];
    const rects: MergeRect[] = [];
    for (const range of this.#merges) {
      const {top, left, bottom, right} = decodeRange(range);
      if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
        merges.push(range);
        continue;
      }
      const [lo, hi] = axis === 'row' ? [top, bottom] : [left, right];
      if (lo >= start && hi < start + count) continue;
      const rect: MergeRect =
        axis === 'row'
          ? {top: shift(top), left, bottom: shift(bottom), right}
          : {top, left: shift(left), bottom, right: shift(right)};
      rects.push(rect);
      merges.push(`${encodeAddress(rect.left, rect.top)}:${encodeAddress(rect.right, rect.bottom)}`);
    }
    this.#merges.length = 0;
    this.#merges.push(...merges);
    this.#mergeRects.length = 0;
    this.#mergeRects.push(...rects);
  }

  // Re-pin the sheet's tables through a splice on the given axis, dropping any table a delete leaves
  // with no row to occupy. `Table` owns the shift arithmetic; the sheet only prunes the casualties.
  #shiftTables(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const survivors = this.#tables.filter(table =>
      axis === 'row' ? table.shiftRows(start, count, delta) : table.shiftColumns(start, count, delta)
    );
    this.#tables.length = 0;
    this.#tables.push(...survivors);
  }

  // Re-pin anchored images through a splice. An anchor point moves like a merge edge: a point before
  // the cut stays, one at or after it shifts by `delta`, and one inside a deleted span clamps to the
  // cut line. Grid points are 0-based, so each is converted to the 1-based coordinate the shared
  // shift arithmetic uses and back. An anchor whose points both move keeps its size; an anchor
  // straddling the cut grows or shrinks, matching how Excel reflows a picture across inserted rows.
  #shiftImages(axis: 'row' | 'col', start: number, count: number, delta: number): void {
    const shift = (v: number): number => (v < start ? v : v >= start + count ? v + delta : start);
    const shiftPoint = (point: AnchorPoint): AnchorPoint => {
      const zeroBased = axis === 'row' ? point.row : point.col;
      const shifted = shift(zeroBased + 1) - 1;
      if (shifted === zeroBased) return point;
      return axis === 'row' ? {...point, row: shifted} : {...point, col: shifted};
    };
    const moved = this.#images.map(image => ({
      imageId: image.imageId,
      anchor: {from: shiftPoint(image.anchor.from), to: shiftPoint(image.anchor.to)},
    }));
    this.#images.length = 0;
    this.#images.push(...moved);
  }

  /**
   * A snapshot of this sheet's transferable content (see {@link WorksheetModel}). Reading it and
   * assigning it onto another sheet — `dst.model = src.model` — clones the source: merges, cells and
   * their styles, column/row metadata, tables, protection, and the page setup all survive, because
   * the getter emits and the setter consumes exactly the same fields. Identity (`name`, `id`) is not
   * part of the model and is never touched by assignment.
   */
  get model(): WorksheetModel {
    const cells: CellModel[] = [];
    for (const [row, cols] of this.#rows) {
      for (const [col, cell] of cols) {
        cells.push({
          row,
          col,
          value: cell.value,
          fill: cell.fill,
          numFmt: cell.numFmt,
          font: cell.font,
          border: cell.border,
          alignment: cell.alignment,
          protection: cell.protection,
          note: cell.note,
        });
      }
    }
    return {
      state: this.state,
      tabColor: this.tabColor,
      properties: {...this.properties},
      outline: {...this.outline},
      pageSetup: {...this.pageSetup},
      pageMargins: {...this.pageMargins},
      headerFooter: {...this.headerFooter},
      columns: [...this.#columns].map(([index, properties]) => ({index, properties: {...properties}})),
      rows: [...this.#rowProperties].map(([number, properties]) => ({number, properties: {...properties}})),
      cells,
      merges: [...this.#merges],
      dataValidations: this.#dataValidations.map(({sqref, rule, extended}) => ({
        sqref,
        rule: cloneDataValidation(rule),
        ...(extended ? {extended: true} : {}),
      })),
      conditionalFormattings: this.#conditionalFormattings.map(cloneConditionalFormatting),
      tables: this.#tables.map(table => table.options),
      protection: this.#protection,
    };
  }

  // Assigning a model replaces this sheet's content wholesale — the sheet becomes the model, with no
  // residue from whatever it held before. Cells are placed at their exact positions (bypassing merge
  // resolution) and merges re-applied after, so a slave's value cannot be misrouted during the load.
  set model(model: WorksheetModel) {
    this.state = model.state;
    this.tabColor = model.tabColor;
    overwrite(this.properties, model.properties);
    overwrite(this.outline, model.outline);
    overwrite(this.pageSetup, model.pageSetup);
    overwrite(this.pageMargins, model.pageMargins);
    overwrite(this.headerFooter, model.headerFooter);

    this.#rows.clear();
    this.#columns.clear();
    this.#rowProperties.clear();
    this.#merges.length = 0;
    this.#mergeRects.length = 0;
    this.#dataValidations.length = 0;
    this.#dataValidationRects.length = 0;
    this.#conditionalFormattings.length = 0;
    this.#tables.length = 0;
    this.#protection = model.protection;

    for (const {index, properties} of model.columns) Object.assign(this.getColumn(index), properties);
    for (const {number, properties} of model.rows) Object.assign(this.getRow(number), properties);
    for (const {row, col, value, fill, numFmt, font, border, alignment, protection, note} of model.cells) {
      const cell = this.#cellAt(row, col);
      cell.value = value;
      cell.fill = fill;
      cell.numFmt = numFmt;
      cell.font = font;
      cell.border = border;
      cell.alignment = alignment;
      cell.protection = protection;
      cell.note = note;
    }
    for (const range of model.merges) this.mergeCells(range);
    for (const {sqref, rule, extended} of model.dataValidations) {
      this.addDataValidation(sqref, rule, extended ? {extended: true} : {});
    }
    for (const formatting of model.conditionalFormattings) this.addConditionalFormatting(formatting);
    for (const options of model.tables) this.addTable(options);
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
    const protection: {flags: SheetProtection['flags']; credential?: SheetProtectionCredential} = {flags};
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
