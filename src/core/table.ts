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
  /** The column's header/display name. Must be unique within the table (Excel disambiguates). */
  readonly name: string;
  /** Literal label shown in the totals row (e.g. `"Total"`), mutually exclusive with a function. */
  readonly totalsRowLabel?: string;
  /** Built-in totals-row aggregate (`"sum"`, `"average"`, `"count"`, …). */
  readonly totalsRowFunction?: string;
}

export interface TableOptions {
  /** Table name — a valid Excel identifier, unique across the workbook. */
  name: string;
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
}

// Excel's table-name grammar: start with a letter, underscore, or backslash; every later
// character a letter, digit, period, or underscore. Unicode letters/digits are allowed.
// Excel additionally forbids a name that *is* a cell reference (`A1`, `R1C1`); we defer
// that rule deliberately — the regression corpus treats cell-reference-shaped names like
// `T1` as valid table names, so enforcing the collision rule here would reject a fixture
// the contract accepts.
const IDENTIFIER = /^[\p{L}\\_][\p{L}\p{N}._]*$/u;

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
  readonly columns: readonly TableColumn[];
  readonly headerRow: boolean;
  readonly totalsRow: boolean;

  readonly #anchorCol: number;
  readonly #anchorRow: number;
  readonly #dataRowCount: number;

  constructor(options: TableOptions) {
    validateTableName(options.name);
    if (options.columns.length === 0) {
      throw new Error(`table "${options.name}" must declare at least one column`);
    }
    if (!Number.isInteger(options.rowCount) || options.rowCount < 0) {
      throw new RangeError(`table "${options.name}" has an invalid data-row count (${options.rowCount})`);
    }
    const {col, row} = decodeAddress(options.ref);
    if (col === undefined || row === undefined) {
      throw new SyntaxError(`table ref "${options.ref}" must anchor at a single cell (e.g. "A1")`);
    }

    this.name = options.name;
    this.columns = options.columns.map(c => ({...c}));
    this.headerRow = options.headerRow ?? true;
    this.totalsRow = options.totalsRow ?? false;
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
   * The options that reconstruct this table — the anchor as a single-cell ref (not the derived
   * full range), the columns, and the data-row count with the header/totals flags. Feeding this
   * back to the constructor yields an equivalent table, so a worksheet model can carry a table
   * losslessly across an export/import round-trip.
   */
  get options(): TableOptions {
    return {
      name: this.name,
      ref: encodeAddress(this.#anchorCol, this.#anchorRow),
      columns: this.columns.map(column => ({...column})),
      rowCount: this.#dataRowCount,
      headerRow: this.headerRow,
      totalsRow: this.totalsRow,
    };
  }

  /** The full A1 range the table occupies: header (if any) + data rows + totals (if any). */
  get ref(): string {
    return `${encodeAddress(this.#anchorCol, this.#anchorRow)}:${encodeAddress(this.#right, this.#bottom)}`;
  }

  /**
   * The autoFilter range — the header row plus the data rows, never the totals row — or
   * `null` for a headerless table, where an autoFilter has nothing to anchor to and Excel
   * treats its presence as corruption.
   */
  get autoFilterRef(): string | null {
    if (!this.headerRow) return null;
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
