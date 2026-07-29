# Row

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Row`

<sub>class</sub>

```ts
class Row {
  readonly number: number;
  get properties(): Readonly<RowProperties> | undefined;
  get height(): number | undefined;
  set height(height: number | undefined);
  get hidden(): boolean | undefined;
  set hidden(hidden: boolean | undefined);
  get outlineLevel(): number | undefined;
  set outlineLevel(outlineLevel: number | undefined);
  get collapsed(): boolean | undefined;
  set collapsed(collapsed: boolean | undefined);
  get fill(): Fill | undefined;
  set fill(fill: Fill | undefined);
  getCell(column: number | string): Cell;
  get cells(): readonly Cell[];
  get values(): (CellValue | undefined)[];
  set values(values: (CellValue | undefined)[]);
}
```

**Members**

- `readonly number: number;` — 1-based row number. Fixed for this handle's lifetime.
- `get properties(): Readonly<RowProperties> | undefined;` — The row's format record if it has one, else `undefined` — a read that never fabricates, so a serializer can ask every row it visits whether there are attributes to emit without giving each one an empty record. Read-only on purpose: `height` and its siblings are how a row is formatted, and they create the record on first write.
- `get height(): number | undefined;` — Row height in points; `undefined` leaves the sheet default in force.
- `get hidden(): boolean | undefined;` — Whether the row is hidden.
- `get outlineLevel(): number | undefined;` — Outline (grouping) depth; 0 or `undefined` means ungrouped.
- `get collapsed(): boolean | undefined;` — Whether this row is the collapsed summary of an outline group.
- `get fill(): Fill | undefined;` — Background fill for the row's cells that carry no fill of their own.
- `getCell(column: number | string): Cell;` — The cell at a column in this row, creating it on first access. The column is a 1-based index (`row.getCell(2)`) or its letters (`row.getCell('B')`). Resolves through merges exactly as `Worksheet.getCell` does: addressing a cell covered by a merged region yields that region's master.
- `get cells(): readonly Cell[];` — The row's materialised cells in ascending column order. Sparse: a column never written to has no cell here, and the array is a fresh snapshot of *which* cells exist — the cells themselves are the live ones.
- `get values(): (CellValue | undefined)[];` — The row's values by position, index 0 being column A. Sparse in the same way `cells` is: a column with no cell is a hole, which is what distinguishes "never written" from a cell holding `null`. Assigning places each value it names and leaves every other column untouched — a hole or an explicit `undefined` skips that column, and a shorter array does not clear the tail. These are `Worksheet.addRow`'s rules, deliberately: `values` is that same row shape addressed by number rather than appended. To *replace* a row, including clearing what it held, splice it — `sheet.spliceRows(n, 1, values)`.
