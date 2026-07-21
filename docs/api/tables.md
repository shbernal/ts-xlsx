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
  get ref(): string;
  get autoFilterRef(): string | null;
  get region(): TableRegion;
}
```

**Members**

- `get rowCount(): number;` — The number of data rows (excludes the header and totals rows). Always defined — a table loaded from a file derives it from the stored range, so reading the height never throws.
- `addRow(values: readonly CellValue[] = []): void;` — Append a data row to the bottom of the table, growing its range by one row and writing `values` left-to-right across its columns. A loaded table exposes its rows the same as a freshly-authored one, so this works identically whether the table was built in memory or read from a file. A table carrying a totals row appends above it: the new data row lands where the totals row sat, and the totals row (with any sheet content below) shifts down by one — exactly what inserting a worksheet row does. That relocation lives in the grid, so a totals-row table not attached to a worksheet throws, as does passing `values` on any detached table — there is nowhere to put them.
- `shiftRows(start: number, count: number, delta: number): boolean;` — Re-pin the table through a row splice: `count` rows removed at the 1-based `start`, then rows inserted so surviving rows below shift by `delta`. A splice entirely above the table moves its whole range by `delta`; one landing inside grows or shrinks the data rows to absorb the change; one that deletes the table's every row removes it. Returns `false` when the table no longer has a row to occupy (the caller drops it), `true` when it survives.
- `shiftColumns(start: number, count: number, delta: number): boolean;` — Re-pin the table through a column splice. A splice entirely to the table's left moves its anchor by `delta`; one to its right leaves it untouched. A splice landing inside the table's columns is structural surgery on named columns with no unambiguous answer, so the table's columns are left as-is (anchor unchanged) rather than fabricated or dropped. Always returns `true`.
- `get options(): TableOptions;` — The options that reconstruct this table — the anchor as a single-cell ref (not the derived full range), the columns, and the data-row count with the header/totals flags. Feeding this back to the constructor yields an equivalent table, so a worksheet model can carry a table losslessly across an export/import round-trip.
- `get ref(): string;` — The full A1 range the table occupies: header (if any) + data rows + totals (if any).
- `get autoFilterRef(): string | null;` — The autoFilter range — the header row plus the data rows, never the totals row — or `null` when the table has no autoFilter: either it is headerless (an autoFilter has nothing to anchor to and Excel treats its presence as corruption) or its `autoFilter` flag is off (a table read without one must not gain one on round-trip).
- `get region(): TableRegion;` — The occupied rectangle, for conflict checks such as overlapping merges.

---

### `TableColumn`

<sub>interface</sub>

One column of a table: a header name and its optional totals-row behaviour.

```ts
interface TableColumn {
    readonly name: string;
    readonly totalsRowLabel?: string;
    readonly totalsRowFunction?: string;
    readonly totalsRowFormula?: string;
    readonly style?: TableColumnStyle;
}
```

---

### `TableOptions`

<sub>interface</sub>

```ts
interface TableOptions {
    name: string;
    displayName?: string;
    ref: string;
    columns: readonly TableColumn[];
    rowCount: number;
    headerRow?: boolean;
    totalsRow?: boolean;
    totalsRowShown?: boolean;
    autoFilter?: boolean;
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
