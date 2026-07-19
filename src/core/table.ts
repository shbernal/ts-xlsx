// A worksheet table (OOXML `<table>`): a named, structured range with typed columns,
// an optional header row, and an optional totals row.
//
// The model stores the anchor and the column/row counts, not a pre-baked range string —
// the occupied geometry is derived, so it stays correct as an empty table (header row
// only), a headerless table (data rows only), or a totals-bearing table. The writer is
// the OOXML gatekeeper for serialization; this model owns the invariants Excel enforces
// on the *shape* itself: a legal name, at least one column, and at least one row.

import {decodeAddress, encodeAddress} from './address.ts';
import type {Alignment, Border, Fill, Font, Protection} from './style.ts';
import type {CellValue} from './value.ts';

/** A per-column cell format applied to a table's body cells — the facets Excel's table-column style
 * bakes into the cells rather than storing as table metadata. Every facet is optional; only the ones
 * set are applied, leaving the rest of each cell's style untouched. */
export interface TableColumnStyle {
  readonly numFmt?: string;
  readonly font?: Partial<Font>;
  readonly fill?: Fill;
  readonly border?: Border;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
}

/** Writes a value into the owning worksheet's grid at a 1-based row/column, applying the column's
 * style (if any) to the cell — the hook a {@link Table} uses to materialise the cells of a row
 * appended through {@link Table.addRow}. A worksheet supplies it when it registers the table; a table
 * built standalone has none and cannot write cell values. */
export type TableCellWriter = (row: number, col: number, value: CellValue, style?: TableColumnStyle) => void;

/** Inserts one empty row into the owning worksheet's grid at a 1-based `row`, shifting that row and
 * everything below it down by one — the hook a {@link Table} with a totals row uses to open a slot
 * for an appended data row above the totals. Relocating the totals row lives in the grid, so a
 * standalone table has no inserter and cannot append past a totals row. */
export type TableRowInserter = (row: number) => void;

/**
 * A table's visual style (`<tableStyleInfo>`): the named style to apply plus the banding/highlight
 * toggles. Every field is a tri-state so a round-trip stays faithful — a value present in the source
 * re-emits, one the source omitted stays omitted rather than being defaulted to `"0"`. A workbook
 * whose part carries no `<tableStyleInfo>` at all leaves {@link TableOptions.style} undefined.
 */
export interface TableStyleInfo {
  /** Named table style to apply (e.g. `"TableStyleMedium2"`, or a workbook-defined custom name). */
  readonly name?: string;
  /** Emphasise the first column. */
  readonly showFirstColumn?: boolean;
  /** Emphasise the last column. */
  readonly showLastColumn?: boolean;
  /** Band the rows (alternating fill). */
  readonly showRowStripes?: boolean;
  /** Band the columns (alternating fill). */
  readonly showColumnStripes?: boolean;
}

/** Copy a style, keeping only its defined fields off the literal so `exactOptionalPropertyTypes`
 * never sees a fabricated `key: undefined` — an absent attribute must stay absent across a copy.
 *
 * The sentinel name `"None"` (Excel's table-style gallery entry for *no* style) is normalised to an
 * absent name: OOXML expresses "unstyled" as a `<tableStyleInfo>` with no `name` attribute, so a
 * literal `name="None"` would reference a style that does not exist and make the file suspect. The
 * banding flags set alongside it are untouched. */
function cloneStyleInfo(style: TableStyleInfo): TableStyleInfo {
  const clone: {-readonly [K in keyof TableStyleInfo]: TableStyleInfo[K]} = {};
  if (style.name !== undefined && style.name !== 'None') clone.name = style.name;
  if (style.showFirstColumn !== undefined) clone.showFirstColumn = style.showFirstColumn;
  if (style.showLastColumn !== undefined) clone.showLastColumn = style.showLastColumn;
  if (style.showRowStripes !== undefined) clone.showRowStripes = style.showRowStripes;
  if (style.showColumnStripes !== undefined) clone.showColumnStripes = style.showColumnStripes;
  return clone;
}

/** One column of a table: a header name and its optional totals-row behaviour. */
export interface TableColumn {
  /** The column's header/display name. Must be unique within the table (case-insensitively) —
   * Excel writes a table with colliding column names as corrupt. A collision supplied at construction
   * is disambiguated deterministically (the first keeps its name, later clashes gain a numeric
   * suffix), the same repair the reader applies to a loaded file, rather than being rejected. */
  readonly name: string;
  /** Literal label shown in the totals row (e.g. `"Total"`), mutually exclusive with a function. */
  readonly totalsRowLabel?: string;
  /** Built-in totals-row aggregate (`"sum"`, `"average"`, `"count"`, …). */
  readonly totalsRowFunction?: string;
  /** A format applied to this column's body cells as they are written (see {@link TableColumnStyle}).
   * Excel bakes a table-column style into the cells rather than storing it as table metadata, so this
   * is an authoring convenience: it round-trips as the affected cells' own styles, not as the table. */
  readonly style?: TableColumnStyle;
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
  /** The table's visual style. Preserved verbatim across a round-trip; when omitted, a freshly
   * authored table is written with Excel's default (`TableStyleMedium2`, banded rows). A part read
   * with no `<tableStyleInfo>` sets this to `undefined`. See {@link TableStyleInfo}. */
  style?: TableStyleInfo;
}

// Excel's table-name grammar: start with a letter, underscore, or backslash; every later
// character a letter, digit, period, or underscore. Unicode letters/digits are allowed.
// Excel additionally forbids a name that *is* a cell reference (`A1`, `R1C1`); we defer
// that rule deliberately — the regression corpus treats cell-reference-shaped names like
// `T1` as valid table names, so enforcing the collision rule here would reject a fixture
// the contract accepts.
const IDENTIFIER = /^[\p{L}\\_][\p{L}\p{N}._]*$/u;

/**
 * Return copies of `columns` with every name made unique (case-insensitively): the first occurrence
 * keeps its name; a later clash gains the smallest numeric suffix that resolves it (`foo`, `foo2`,
 * `foo3`, …). OOXML requires unique column names within a table — Excel treats a collision as
 * corruption — so this is applied both when a table is authored and when one is read from a file,
 * keeping the two paths identical rather than rejecting a name list the reader would accept.
 */
export function disambiguateColumnNames(columns: readonly TableColumn[]): TableColumn[] {
  const seen = new Set<string>();
  return columns.map(column => {
    let candidate = column.name;
    for (let n = 2; seen.has(candidate.toLowerCase()); n++) candidate = `${column.name}${n}`;
    seen.add(candidate.toLowerCase());
    return candidate === column.name ? {...column} : {...column, name: candidate};
  });
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
  readonly style: TableStyleInfo | undefined;

  // The anchor and data-row count move when a row/column splice shifts or resizes the table, so
  // they are mutable behind the class's controlled `shiftRows`/`shiftColumns` methods.
  #anchorCol: number;
  #anchorRow: number;
  #dataRowCount: number;

  // Set by the worksheet that registers this table so an appended row can be written into the grid.
  // A table constructed standalone (a unit test, a bare model) has none — appending values then
  // throws rather than silently dropping them.
  readonly #writeCell: TableCellWriter | undefined;

  // Supplied alongside #writeCell by the registering worksheet. A totals-row table appends by
  // inserting a grid row above the totals; a standalone table has neither hook.
  readonly #insertRow: TableRowInserter | undefined;

  constructor(options: TableOptions, writeCell?: TableCellWriter, insertRow?: TableRowInserter) {
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
    this.displayName = options.displayName ?? options.name;
    this.columns = disambiguateColumnNames(options.columns);
    this.headerRow = options.headerRow ?? true;
    this.totalsRow = options.totalsRow ?? false;
    this.totalsRowShown = options.totalsRowShown;
    this.style = options.style === undefined ? undefined : cloneStyleInfo(options.style);
    // A header table gains an autoFilter by default (Excel's behaviour when a table is inserted);
    // a headerless table can never carry one — an autoFilter has no header row to anchor to.
    this.autoFilter = this.headerRow && (options.autoFilter ?? true);
    this.#anchorCol = col;
    this.#anchorRow = row;
    this.#dataRowCount = options.rowCount;
    this.#writeCell = writeCell;
    this.#insertRow = insertRow;

    if (this.#rowSpan < 1) {
      throw new Error(`table "${this.name}" has no rows — it needs a header row or at least one data row`);
    }
  }

  get columnCount(): number {
    return this.columns.length;
  }

  /** The number of data rows (excludes the header and totals rows). Always defined — a table loaded
   * from a file derives it from the stored range, so reading the height never throws. */
  get rowCount(): number {
    return this.#dataRowCount;
  }

  /**
   * Append a data row to the bottom of the table, growing its range by one row and writing `values`
   * left-to-right across its columns. A loaded table exposes its rows the same as a freshly-authored
   * one, so this works identically whether the table was built in memory or read from a file.
   *
   * A table carrying a totals row appends above it: the new data row lands where the totals row sat,
   * and the totals row (with any sheet content below) shifts down by one — exactly what inserting a
   * worksheet row does. That relocation lives in the grid, so a totals-row table not attached to a
   * worksheet throws, as does passing `values` on any detached table — there is nowhere to put them.
   */
  addRow(values: readonly CellValue[] = []): void {
    if (values.length > this.columnCount) {
      throw new RangeError(
        `row has ${values.length} values but table "${this.name}" has ${this.columnCount} columns`
      );
    }

    // The append point is the row directly below the last data row: the totals row when one exists,
    // otherwise the first free row under the table.
    const target = this.#anchorRow + (this.headerRow ? 1 : 0) + this.#dataRowCount;

    if (this.totalsRow) {
      if (this.#insertRow === undefined) {
        throw new Error(
          `table "${this.name}" is not attached to a worksheet — cannot relocate its totals row to append a data row`
        );
      }
      // Opening a grid slot at the totals row shifts the totals down and grows this table by one
      // through the sheet's own table re-pinning, so #dataRowCount is not bumped again here.
      this.#insertRow(target);
      values.forEach((value, index) =>
        this.#writeCell?.(target, this.#anchorCol + index, value, this.columns[index]?.style)
      );
      return;
    }

    if (values.length > 0) {
      if (this.#writeCell === undefined) {
        throw new Error(
          `table "${this.name}" is not attached to a worksheet — cannot write appended row values`
        );
      }
      values.forEach((value, index) =>
        this.#writeCell?.(target, this.#anchorCol + index, value, this.columns[index]?.style)
      );
    }
    this.#dataRowCount += 1;
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
    if (this.style !== undefined) options.style = cloneStyleInfo(this.style);
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
