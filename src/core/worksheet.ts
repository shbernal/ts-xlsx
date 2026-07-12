// A worksheet: a sparsely-populated grid of cells addressed by A1 reference.
//
// Storage is sparse by construction — a spreadsheet is mostly empty, so cells
// materialise on first access and only occupied positions cost memory. Column and
// row metadata (widths, heights, visibility, outline grouping) are stored apart from
// the cell grid, because a column or row can carry formatting while holding no cells.
// Merges and views layer on in later slices.

import {decodeAddress} from './address.ts';
import {Cell} from './cell.ts';
import {Table, type TableOptions} from './table.ts';

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

/** Per-column formatting. A column may exist purely to carry these, with no cells. */
export interface ColumnProperties {
  /** Column width in character units. */
  width?: number;
  /** Whether the column is hidden. */
  hidden?: boolean;
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
}

export class Worksheet {
  readonly name: string;
  /** 1-based workbook-assigned id, stable for the sheet's lifetime. */
  readonly id: number;
  state: WorksheetState['state'];

  /** Sheet-level format defaults. Mutate in place: `sheet.properties.defaultRowHeight = 20`. */
  readonly properties: WorksheetProperties = {};

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
  // Tables and merged ranges are sheet-level overlays on the grid, not cell storage.
  readonly #tables: Table[] = [];
  readonly #merges: string[] = [];

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
   * @throws {SyntaxError} if the reference does not resolve to a single cell.
   */
  getCell(reference: string): Cell {
    const {col, row} = decodeAddress(reference);
    if (col === undefined || row === undefined) {
      throw new SyntaxError(`"${reference}" is not a single-cell reference — it omits a column or row`);
    }
    return this.#cellAt(row, col);
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

  /** Merge a range of cells (`"A1:B2"`). Overlap validation happens at write time. */
  mergeCells(range: string): void {
    this.#merges.push(range);
  }

  /** The merged ranges on this sheet, in the order they were added. */
  get merges(): readonly string[] {
    return this.#merges;
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
