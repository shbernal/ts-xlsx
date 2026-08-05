# Tables

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Table`

<sub>class</sub>

```ts
class Table {
  readonly name: string;
  readonly displayName: string;
  readonly columns: readonly TableColumn[];
  readonly headerRow: boolean;
  readonly totalsRow: boolean;
  readonly totalsRowShown: boolean | undefined;
  readonly autoFilter: boolean;
  readonly style: TableStyleInfo | undefined;
  get columnCount(): number;
  get rowCount(): number;
  addRow(values: readonly CellValue[] = []): void;
  shiftRows(start: number, count: number, delta: number): boolean;
  shiftColumns(start: number, count: number, delta: number): boolean;
  get options(): TableOptions;
  get range(): string;
  get autoFilterRef(): string | undefined;
  get region(): TableRegion;
}
```

**Members**

- `get rowCount(): number;` — The number of data rows (excludes the header and totals rows). Always defined — a table loaded from a file derives it from the stored range, so reading the height never throws.
- `addRow(values: readonly CellValue[] = []): void;` — Append a data row to the bottom of the table, growing its range by one row and writing `values` left-to-right across its columns. A loaded table exposes its rows the same as a freshly-authored one, so this works identically whether the table was built in memory or read from a file. A table carrying a totals row appends above it: the new data row lands where the totals row sat, and the totals row (with any sheet content below) shifts down by one — exactly what inserting a worksheet row does. That relocation lives in the grid, so a totals-row table not attached to a worksheet throws, as does passing `values` on any detached table — there is nowhere to put them.
- `shiftRows(start: number, count: number, delta: number): boolean;` — Re-pin the table through a row splice: `count` rows removed at the 1-based `start`, then rows inserted so surviving rows below shift by `delta`. A splice entirely above the table moves its whole range by `delta`; one landing inside grows or shrinks the data rows to absorb the change; one that deletes the table's every row removes it. Returns `false` when the table no longer has a row to occupy (the caller drops it), `true` when it survives.
- `shiftColumns(start: number, count: number, delta: number): boolean;` — Re-pin the table through a column splice. A splice entirely to the table's left moves its anchor by `delta`; one to its right leaves it untouched. A splice landing inside the table's columns is structural surgery on named columns with no unambiguous answer, so the table's columns are left as-is (anchor unchanged) rather than fabricated or dropped. Always returns `true`.
- `get options(): TableOptions;` — The options that reconstruct this table — the anchor as a single-cell ref (not the derived full range), the columns, and the data-row count with the header/totals flags. Feeding this back to the constructor yields an equivalent table, so a worksheet model can carry a table losslessly across an export/import round-trip.
- `get range(): string;` — The full A1 range the table occupies: header (if any) + data rows + totals (if any). Distinct from `TableOptions.ref` (and `options`'s own `ref`), which is only the single-cell anchor a table is constructed from — this is the anchor plus the columns/rows it has grown to cover.
- `get autoFilterRef(): string | undefined;` — The autoFilter range — the header row plus the data rows, never the totals row — or `undefined` when the table has no autoFilter: either it is headerless (an autoFilter has nothing to anchor to and Excel treats its presence as corruption) or its `autoFilter` flag is off (a table read without one must not gain one on round-trip).
- `get region(): TableRegion;` — The occupied rectangle, for conflict checks such as overlapping merges.

---

### `TableColumn`

<sub>interface</sub>

One column of a table: a header name and its optional totals-row behaviour.

```ts
interface TableColumn {
  /** The column's header/display name. Must be unique within the table (case-insensitively) —
   * Excel writes a table with colliding column names as corrupt. A collision supplied at construction
   * is disambiguated deterministically (the first keeps its name, later clashes gain a numeric
   * suffix), the same repair the reader applies to a loaded file, rather than being rejected. */
  readonly name: string;
  /** Literal label shown in the totals row (e.g. `"Total"`), mutually exclusive with a function. */
  readonly totalsRowLabel?: string;
  /** Built-in totals-row aggregate (`"sum"`, `"average"`, `"count"`, …), or `"custom"` when the
   * column's total is the arbitrary formula in {@link totalsRowFormula} rather than a `SUBTOTAL`. */
  readonly totalsRowFunction?: TotalsRowFunction;
  /** The formula (no leading `=`) backing a `totalsRowFunction: "custom"` column — OOXML's
   * `<totalsRowFormula>` child. Round-tripped verbatim and written into the totals cell as the
   * cell's formula. Meaningful only alongside `totalsRowFunction: "custom"`; ignored otherwise. */
  readonly totalsRowFormula?: string;
  /** A format applied to this column's body cells as they are written (see {@link TableColumnStyle}).
   * Excel bakes a table-column style into the cells rather than storing it as table metadata, so this
   * is an authoring convenience: it round-trips as the affected cells' own styles, not as the table. */
  readonly style?: TableColumnStyle;
}
```

---

### `TableColumnStyle`

<sub>type</sub>

A per-column cell format applied to a table's body cells — the facets Excel's table-column style
bakes into the cells rather than storing as table metadata. Every facet (`CellStyle`) is
optional; only the ones set are applied, leaving the rest of each cell's style untouched.

```ts
type TableColumnStyle = Readonly<CellStyle>;
```

---

### `TableOptions`

<sub>interface</sub>

```ts
interface TableOptions {
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
```

---

### `TableRegion`

<sub>interface</sub>

The rectangle a table occupies, in 1-based coordinates.

```ts
interface TableRegion {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}
```

---

### `TableStyleInfo`

<sub>interface</sub>

A table's visual style (`<tableStyleInfo>`): the named style to apply plus the banding/highlight
toggles. Every field is a tri-state so a round-trip stays faithful — a value present in the source
re-emits, one the source omitted stays omitted rather than being defaulted to `"0"`. A workbook
whose part carries no `<tableStyleInfo>` at all leaves `TableOptions.style` undefined.

```ts
interface TableStyleInfo {
  /**
   * Named table style to apply — one of the built-in gallery (`"TableStyleMedium2"`, …) or a custom
   * one the workbook defines with {@link Workbook.addTableStyle}.
   *
   * **Not validated.** A name that matches nothing renders the table unstyled, silently — but this
   * library must not be the thing that rejects it. A reader has to accept a name from a newer Excel
   * than the gallery list it was built with, and a writer that threw would make round-tripping such a
   * file impossible; there is also no diagnostics channel to warn through, so the only options were
   * "throw" and "accept". Accepting is the one that never makes a readable file unreadable. If a
   * warning channel is ever added, this is the first thing that should use it.
   */
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
```
