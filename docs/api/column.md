# Column

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Column`

<sub>class</sub>

```ts
class Column {
  readonly index: number;
  get letter(): string;
  get properties(): Readonly<ColumnProperties> | undefined;
  get key(): string | undefined;
  set key(key: string | undefined);
  get width(): number | undefined;
  set width(width: number | undefined);
  get hidden(): boolean | undefined;
  set hidden(hidden: boolean | undefined);
  get outlineLevel(): number | undefined;
  set outlineLevel(outlineLevel: number | undefined);
  get collapsed(): boolean | undefined;
  set collapsed(collapsed: boolean | undefined);
  get fill(): Fill | undefined;
  set fill(fill: Fill | undefined);
  get numFmt(): string | undefined;
  set numFmt(numFmt: string | undefined);
  get font(): Font | undefined;
  set font(font: Font | undefined);
  get border(): Border | undefined;
  set border(border: Border | undefined);
  get alignment(): Alignment | undefined;
  set alignment(alignment: Alignment | undefined);
  get protection(): Protection | undefined;
  set protection(protection: Protection | undefined);
  getCell(row: number): Cell;
  get cells(): readonly Cell[];
  get values(): (CellValue | undefined)[];
  set values(values: (CellValue | undefined)[]);
}
```

**Members**

- `readonly index: number;` — 1-based column index. Fixed for this handle's lifetime.
- `get letter(): string;` — The column's letters (`"B"`) — the spreadsheet-facing name for `index`.
- `get properties(): Readonly<ColumnProperties> | undefined;` — The column's format record if it has one, else `undefined` — a read that never fabricates, so a serializer can ask every column it visits whether there are attributes to emit without giving each one an empty record. Read-only on purpose: `width` and its siblings are how a column is formatted, and they create the record on first write.
- `get key(): string | undefined;` — Stable key naming this column so a keyed-object row (see `Worksheet.addRow`) can place a value under it by name rather than position. In-memory only — never serialized to OOXML.
- `get width(): number | undefined;` — Column width in character units; `undefined` leaves the sheet default in force.
- `get hidden(): boolean | undefined;` — Whether the column is hidden.
- `get outlineLevel(): number | undefined;` — Outline (grouping) depth; 0 or `undefined` means ungrouped.
- `get collapsed(): boolean | undefined;` — Whether this column is the collapsed summary of an outline group.
- `get fill(): Fill | undefined;` — Default fill for the column's cells that set none of their own.
- `get numFmt(): string | undefined;` — Default number format for the column's cells that set none of their own.
- `get font(): Font | undefined;` — Default font for the column's cells that set none of their own.
- `get border(): Border | undefined;` — Default border for the column's cells that set none of their own.
- `get alignment(): Alignment | undefined;` — Default alignment for the column's cells that set none of their own.
- `get protection(): Protection | undefined;` — Default protection flags for the column's cells that set none of their own.
- `getCell(row: number): Cell;` — The cell at a 1-based row number in this column, creating it on first access. Resolves through merges exactly as `Worksheet.getCell` does.
- `get cells(): readonly Cell[];` — The column's materialised cells in ascending row order. Sparse: a row that never wrote to this column has no cell here.
- `get values(): (CellValue | undefined)[];` — The column's values by position, index 0 being row 1. Sparse in the same way `cells` is: a row with no cell in this column is a hole, which is what distinguishes "never written" from a cell holding `null`. Assigning places each value it names and leaves every other row untouched, mirroring `Row.values` — a hole or an explicit `undefined` skips that row, and a shorter array does not clear the tail.
