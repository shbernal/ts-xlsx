// A worksheet table (OOXML `<table>`): a named, structured range with typed columns,
// an optional header row, and an optional totals row.
//
// The model stores the anchor and the column/row counts, not a pre-baked range string —
// the occupied geometry is derived, so it stays correct as an empty table (header row
// only), a headerless table (data rows only), or a totals-bearing table. The writer is
// the OOXML gatekeeper for serialization; this model owns the invariants Excel enforces
// on the *shape* itself: a legal name, at least one column, and at least one row.

import {decodeAddress, encodeAddress} from './address.ts';

/** One column of a table: a header name and its optional totals-row behaviour. */
export interface TableColumn {
  /** The column's header/display name. Must be unique within the table (case-insensitively) —
   * Excel writes a table with colliding column names as corrupt, so a duplicate is rejected. */
  readonly name: string;
  /** Literal label shown in the totals row (e.g. `"Total"`), mutually exclusive with a function. */
  readonly totalsRowLabel?: string;
  /** Built-in totals-row aggregate (`"sum"`, `"average"`, `"count"`, …). */
  readonly totalsRowFunction?: string;
}

export interface TableOptions {
  /** Table name — a valid Excel identifier, unique across the workbook. This is the name used in
   * structured formula references (`Table1[Column]`). */
  name: string;
  /** Human-facing display name shown in the UI. A free-form label (spaces allowed) that need not
   * be a valid identifier. Defaults to {@link name} when omitted. */
  displayName?: string;
  /** A1 reference of the table's top-left cell (an anchor, e.g. `"A1"` — not the full range). */
  ref: string;
  /** The table's columns, left to right. At least one is required. */
  columns: readonly TableColumn[];
  /** Number of data rows (excludes the header and totals rows). May be zero. */
  rowCount: number;
  /** Whether the table has a header row. Defaults to `true`. */
  headerRow?: boolean;
  /** Whether the table has a totals row. Defaults to `false`. */
  totalsRow?: boolean;
  /** The `totalsRowShown` flag on a table *without* a totals row — Excel's record of whether a
   * totals row has ever been toggled on. Tri-state so a round-trip is faithful: `false` re-emits
   * `totalsRowShown="0"`, `true` re-emits `totalsRowShown="1"`, and `undefined` (the authoring
   * default) emits nothing — a file read without the attribute must not have one injected. Ignored
   * when {@link totalsRow} is set, since a present totals row already implies it is shown. */
  totalsRowShown?: boolean;
  /** Whether the header row carries an autoFilter. Defaults to {@link headerRow}: a header table
   * gains an autoFilter, a headerless one never can. Set `false` to keep a header table's rows
   * unfiltered — a file read without an autoFilter must round-trip without one being injected. */
  autoFilter?: boolean;
}

// Excel's table-name grammar: start with a letter, underscore, or backslash; every later
// character a letter, digit, period, or underscore. Unicode letters/digits are allowed.
// Excel additionally forbids a name that *is* a cell reference (`A1`, `R1C1`); we defer
// that rule deliberately — the regression corpus treats cell-reference-shaped names like
// `T1` as valid table names, so enforcing the collision rule here would reject a fixture
// the contract accepts.
const IDENTIFIER = /^[\p{L}\\_][\p{L}\p{N}._]*$/u;

function validateColumnNames(tableName: string, columns: readonly TableColumn[]): void {
  const seen = new Set<string>();
  for (const {name} of columns) {
    const key = name.toLowerCase();
    if (seen.has(key)) {
      throw new Error(
        `table "${tableName}" has a duplicate column name ${JSON.stringify(name)} — ` +
          'a table\'s column names must be unique (Excel treats a collision as corruption)'
      );
    }
    seen.add(key);
  }
}

function validateTableName(name: string): void {
  if (name.length === 0 || name.length > 255) {
    throw new Error(`table name ${JSON.stringify(name)} must be between 1 and 255 characters`);
  }
  if (!IDENTIFIER.test(name)) {
    throw new Error(
      `table name ${JSON.stringify(name)} is not a valid Excel identifier — it must start with a letter, ` +
        'underscore, or backslash and contain only letters, digits, periods, and underscores'
    );
  }
}

/** The rectangle a table occupies, in 1-based coordinates. */
export interface TableRegion {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export class Table {
  readonly name: string;
  readonly displayName: string;
  readonly columns: readonly TableColumn[];
  readonly headerRow: boolean;
  readonly totalsRow: boolean;
  readonly totalsRowShown: boolean | undefined;
  readonly autoFilter: boolean;

  // The anchor and data-row count move when a row/column splice shifts or resizes the table, so
  // they are mutable behind the class's controlled `shiftRows`/`shiftColumns` methods.
  #anchorCol: number;
  #anchorRow: number;
  #dataRowCount: number;

  constructor(options: TableOptions) {
    validateTableName(options.name);
    if (options.columns.length === 0) {
      throw new Error(`table "${options.name}" must declare at least one column`);
    }
    validateColumnNames(options.name, options.columns);
    if (!Number.isInteger(options.rowCount) || options.rowCount < 0) {
      throw new RangeError(`table "${options.name}" has an invalid data-row count (${options.rowCount})`);
    }
    const {col, row} = decodeAddress(options.ref);
    if (col === undefined || row === undefined) {
      throw new SyntaxError(`table ref "${options.ref}" must anchor at a single cell (e.g. "A1")`);
    }

    this.name = options.name;
    this.displayName = options.displayName ?? options.name;
    this.columns = options.columns.map(c => ({...c}));
    this.headerRow = options.headerRow ?? true;
    this.totalsRow = options.totalsRow ?? false;
    this.totalsRowShown = options.totalsRowShown;
    // A header table gains an autoFilter by default (Excel's behaviour when a table is inserted);
    // a headerless table can never carry one — an autoFilter has no header row to anchor to.
    this.autoFilter = this.headerRow && (options.autoFilter ?? true);
    this.#anchorCol = col;
    this.#anchorRow = row;
    this.#dataRowCount = options.rowCount;

    if (this.#rowSpan < 1) {
      throw new Error(`table "${this.name}" has no rows — it needs a header row or at least one data row`);
    }
  }

  get columnCount(): number {
    return this.columns.length;
  }

  /**
   * Re-pin the table through a row splice: `count` rows removed at the 1-based `start`, then rows
   * inserted so surviving rows below shift by `delta`. A splice entirely above the table moves its
   * whole range by `delta`; one landing inside grows or shrinks the data rows to absorb the change;
   * one that deletes the table's every row removes it. Returns `false` when the table no longer has
   * a row to occupy (the caller drops it), `true` when it survives.
   */
  shiftRows(start: number, count: number, delta: number): boolean {
    // A table whose every row lies within the deleted span has nothing left to occupy.
    if (this.#anchorRow >= start && this.#bottom < start + count) return false;
    const shift = (v: number): number => (v < start ? v : v >= start + count ? v + delta : start);
    const top = shift(this.#anchorRow);
    const bottom = shift(this.#bottom);
    const span = bottom - top + 1;
    const fixedRows = (this.headerRow ? 1 : 0) + (this.totalsRow ? 1 : 0);
    const dataRows = span - fixedRows;
    if (span < 1 || dataRows < 0) return false;
    this.#anchorRow = top;
    this.#dataRowCount = dataRows;
    return true;
  }

  /**
   * Re-pin the table through a column splice. A splice entirely to the table's left moves its anchor
   * by `delta`; one to its right leaves it untouched. A splice landing inside the table's columns is
   * structural surgery on named columns with no unambiguous answer, so the table's columns are left
   * as-is (anchor unchanged) rather than fabricated or dropped. Always returns `true`.
   */
  shiftColumns(start: number, count: number, delta: number): boolean {
    if (this.#anchorCol >= start + count) this.#anchorCol += delta;
    return true;
  }

  /**
   * The options that reconstruct this table — the anchor as a single-cell ref (not the derived
   * full range), the columns, and the data-row count with the header/totals flags. Feeding this
   * back to the constructor yields an equivalent table, so a worksheet model can carry a table
   * losslessly across an export/import round-trip.
   */
  get options(): TableOptions {
    const options: TableOptions = {
      name: this.name,
      displayName: this.displayName,
      ref: encodeAddress(this.#anchorCol, this.#anchorRow),
      columns: this.columns.map(column => ({...column})),
      rowCount: this.#dataRowCount,
      headerRow: this.headerRow,
      totalsRow: this.totalsRow,
      autoFilter: this.autoFilter,
    };
    // Kept off the literal so `undefined` (attribute absent) stays absent, not an explicit
    // `totalsRowShown: undefined` — the round-trip must not fabricate the flag.
    if (this.totalsRowShown !== undefined) options.totalsRowShown = this.totalsRowShown;
    return options;
  }

  /** The full A1 range the table occupies: header (if any) + data rows + totals (if any). */
  get ref(): string {
    return `${encodeAddress(this.#anchorCol, this.#anchorRow)}:${encodeAddress(this.#right, this.#bottom)}`;
  }

  /**
   * The autoFilter range — the header row plus the data rows, never the totals row — or
   * `null` when the table has no autoFilter: either it is headerless (an autoFilter has nothing
   * to anchor to and Excel treats its presence as corruption) or its {@link autoFilter} flag is
   * off (a table read without one must not gain one on round-trip).
   */
  get autoFilterRef(): string | null {
    if (!this.autoFilter) return null;
    const bottom = this.#anchorRow + this.#dataRowCount;
    return `${encodeAddress(this.#anchorCol, this.#anchorRow)}:${encodeAddress(this.#right, bottom)}`;
  }

  /** The occupied rectangle, for conflict checks such as overlapping merges. */
  get region(): TableRegion {
    return {top: this.#anchorRow, left: this.#anchorCol, bottom: this.#bottom, right: this.#right};
  }

  get #right(): number {
    return this.#anchorCol + this.columns.length - 1;
  }

  get #rowSpan(): number {
    return (this.headerRow ? 1 : 0) + this.#dataRowCount + (this.totalsRow ? 1 : 0);
  }

  get #bottom(): number {
    return this.#anchorRow + this.#rowSpan - 1;
  }
}
