# Worksheet

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellModel`

<sub>interface</sub>

One materialised cell in a `WorksheetModel`: its position, value, and per-cell style facets.

```ts
interface CellModel extends CellStyle {
    readonly row: number;
    readonly col: number;
    value: CellValue;
    note?: string | undefined;
}
```

---

### `ColumnProperties`

<sub>interface</sub>

Per-column formatting. A column may exist purely to carry these, with no cells. The style
facets are *defaults* for the column's cells: a cell that sets a facet of its own wins, but
one that leaves a facet unset inherits the column's — the same precedence Excel applies, and
symmetric with how a `RowProperties` fill defaults a row's cells.

```ts
interface ColumnProperties extends CellStyle {
    key?: string;
    width?: number;
    hidden?: boolean;
    outlineLevel?: number;
    collapsed?: boolean;
}
```

---

### `OutlineProperties`

<sub>interface</sub>

Placement of an outline's summary rows/columns. Excel's defaults are summary *below* the detail
rows and *right* of the detail columns; setting either to `false` inverts that placement so an
author who groups upward gets a file that honours it. An unset flag is omitted from the written
`<outlinePr>`, and an empty object emits no `<outlinePr>` at all.

```ts
interface OutlineProperties {
    summaryBelow?: boolean;
    summaryRight?: boolean;
}
```

---

### `RowInput`

<sub>type</sub>

A row handed to `Worksheet.addRow`: a positional array of cell values (a hole or `undefined`
leaves that column untouched), or an object keyed by column `ColumnProperties.key` whose
values land under the matching columns.

```ts
type RowInput = (CellValue | undefined)[] | Record<string, CellValue>;
```

---

### `RowProperties`

<sub>interface</sub>

Per-row formatting. A row may exist purely to carry these, with no cells.

```ts
interface RowProperties {
    height?: number;
    hidden?: boolean;
    outlineLevel?: number;
    collapsed?: boolean;
    fill?: Fill;
}
```

---

### `SheetView`

<sub>interface</sub>

A worksheet's frozen-pane view. `state` `'frozen'` locks the top `ySplit` rows and left `xSplit`
columns in place while the rest scrolls; `'normal'` (the default) has no split and emits no
`<pane>` — writing a normal view leaves no leftover pane markup that would trip Excel's repair
prompt. An empty object is a normal view.

```ts
interface SheetView {
    state?: 'normal' | 'frozen';
    xSplit?: number;
    ySplit?: number;
    topLeftCell?: string;
}
```

---

### `Worksheet`

<sub>class</sub>

```ts
class Worksheet {
  readonly name: string;
  readonly id: number;
  state: WorksheetState['state'];
  tabColor: Color | undefined;
  readonly properties: WorksheetProperties = {};
  readonly outline: OutlineProperties = {};
  readonly view: SheetView = {};
  readonly pageSetup: PageSetup = {};
  readonly printOptions: PrintOptions = {};
  readonly pageMargins: PageMargins = {};
  readonly headerFooter: HeaderFooter = {};
  readonly rowBreaks: PageBreak[] = [];
  readonly columnBreaks: PageBreak[] = [];
  getCell(reference: string): Cell;
  hasCell(row: number, col: number): boolean;
  rowProperties(number: number): RowProperties | undefined;
  columnProperties(index: number): ColumnProperties | undefined;
  getColumn(index: number): ColumnProperties;
  getRow(number: number): RowProperties;
  get rowCount(): number;
  get actualRowCount(): number;
  get columnCount(): number;
  *columns(): IterableIterator<{
    readonly index: number;
    readonly properties: ColumnProperties;
}>;
  *rows(): IterableIterator<{
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
}>;
  addTable(options: TableOptions): Table;
  get tables(): readonly Table[];
  getTable(name: string): Table | undefined;
  addPivotTable(options: PivotTableOptions): PivotTable;
  get pivotTables(): readonly PivotTable[];
  get loadedPivotTables(): readonly ParsedPivotTable[];
  addCommentThread(thread: CommentThread): void;
  get commentThreads(): readonly CommentThread[];
  commentThreadAt(reference: string): CommentThread | undefined;
  addImage(imageId: number, anchor: {
    readonly tl: AnchorPoint;
    readonly br: AnchorPoint;
    readonly editAs?: ImageEditAs;
}): void;
  addImage(imageId: number, anchor: {
    readonly tl: AnchorPoint;
    readonly ext: {
        readonly width: number;
        readonly height: number;
    };
}): void;
  addImage(imageId: number, anchor: {
    readonly tl: AnchorPoint;
    readonly br: AnchorPoint;
    readonly editAs?: ImageEditAs;
} | {
    readonly tl: AnchorPoint;
    readonly ext: {
        readonly width: number;
        readonly height: number;
    };
}): void;
  addImageAnchor(imageId: number, anchor: ImageAnchor): void;
  removeImage(imageId: number): void;
  get images(): readonly AnchoredImage[];
  addBackgroundImage(imageId: number): void;
  removeBackgroundImage(): void;
  get backgroundImageId(): number | undefined;
  get preservedReferences(): readonly PreservedWorksheetReference[];
  mergeCells(range: string): void;
  get merges(): readonly string[];
  get autoFilter(): AutoFilter | undefined;
  set autoFilter(filter: string | AutoFilter | undefined);
  unmergeCells(range: string): boolean;
  addDataValidation(sqref: string, rule: DataValidation, options: {
    extended?: boolean;
} = {}): void;
  get dataValidations(): readonly DataValidationEntry[];
  addConditionalFormatting(formatting: ConditionalFormatting): void;
  get conditionalFormattings(): readonly ConditionalFormatting[];
  dataValidationAt(reference: string): DataValidation | undefined;
  spliceRows(start: number, count: number, ...inserts: RowInput[]): void;
  insertRow(pos: number, values: RowInput): void;
  addRow(values: RowInput): Cell[];
  addRows(rows: RowInput[]): Cell[][];
  freeze(ySplit = 1, xSplit = 0): void;
  unfreeze(): void;
  duplicateRow(start: number, options: {
    count?: number;
    insert?: boolean;
} = {}): void;
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void;
  insertColumn(pos: number, values: CellValue[]): void;
  addColumn(values: CellValue[]): Cell[];
  addColumns(columns: CellValue[][]): Cell[][];
  get model(): WorksheetModel;
  set model(model: WorksheetModel);
  protect(password?: string, options: SheetProtectionOptions = {}): void;
  unprotect(): void;
  get protection(): SheetProtection | undefined;
}
```

**Members**

- `readonly id: number;` — 1-based workbook-assigned id, stable for the sheet's lifetime.
- `tabColor: Color | undefined;` — Colour of the sheet's tab, as an ARGB/theme `Color`. `undefined` leaves the tab its default colour; the writer emits no `<tabColor>` for an uncoloured sheet, so a round-trip never fabricates one.
- `readonly properties: WorksheetProperties = {};` — Sheet-level format defaults. Mutate in place: `sheet.properties.defaultRowHeight = 20`.
- `readonly outline: OutlineProperties = {};` — Outline summary-position flags. Mutate in place: `sheet.outline.summaryBelow = false`. Empty means unset — the writer emits no `<outlinePr>` and a round-trip never fabricates one.
- `readonly view: SheetView = {};` — The sheet's frozen-pane view. Empty (a normal view) emits no `<pane>`. Use `freeze` and `unfreeze` for the common cases, or mutate in place for finer control.
- `readonly pageSetup: PageSetup = {};` — Print-scaling and orientation. Mutate in place: `sheet.pageSetup.fitToPage = true`. Empty means unset — the writer emits neither `<pageSetUpPr>` nor `<pageSetup>` and a round-trip never fabricates them.
- `readonly printOptions: PrintOptions = {};` — Print-toggle flags (`<printOptions>`): centring, and whether headings/gridlines print. Mutate in place: `sheet.printOptions.gridLines = true`. Empty means unset — the writer emits no element and a round-trip never fabricates one.
- `readonly pageMargins: PageMargins = {};` — Print margins. Mutate in place: `sheet.pageMargins.left = 0.5`. Empty means unset.
- `readonly headerFooter: HeaderFooter = {};` — Page header/footer text. Mutate in place: `sheet.headerFooter.oddHeader = '&C&"..."'`.
- `readonly rowBreaks: PageBreak[] = [];` — Manual horizontal page breaks (`<rowBreaks>`): each break's `id` is a row the print layout splits before. Mutate in place: `sheet.rowBreaks.push({id: 3})`. Empty means no manual row breaks and the writer emits no `<rowBreaks>` element.
- `readonly columnBreaks: PageBreak[] = [];` — Manual vertical page breaks (`<colBreaks>`): each break's `id` is a column the print layout splits before. Mutate in place: `sheet.columnBreaks.push({id: 3})`. Empty means no manual column breaks and the writer emits no `<colBreaks>` element.
- `getCell(reference: string): Cell;` — Get the cell at an A1 reference, creating it on first access. The reference must name both a column and a row (`"B3"`); a whole-row or whole-column reference is not a cell and is rejected. Addressing a cell covered by a merged region resolves to that region's master (top-left) cell, mirroring how a spreadsheet treats the merge as one cell: a value or style written through a covered address lands on the master, and reading a covered address returns the master's. Only the master ever holds an independent value, so the serialized sheet stays well-formed (no stray value on a covered cell).
- `hasCell(row: number, col: number): boolean;` — Whether a cell has been materialised at the given 1-based position.
- `rowProperties(number: number): RowProperties | undefined;` — The format properties for a 1-based row number if any were set, or `undefined` — a read-only peek that never creates a record, so a serializer can render a row's attributes without fabricating an empty one for every row it visits. Use `getRow` to create-on-access.
- `columnProperties(index: number): ColumnProperties | undefined;` — The format properties for a 1-based column index if any were set, or `undefined` — a read-only peek that never creates a record, so a serializer can render a column's attributes without fabricating an empty one for every column it visits. Use `getColumn` to create-on-access.
- `getColumn(index: number): ColumnProperties;` — Get the mutable format properties for a 1-based column index, creating the record on first access. Setting properties here does not materialise any cells.
- `getRow(number: number): RowProperties;` — Get the mutable format properties for a 1-based row number, creating the record on first access. This is row *metadata* (height, visibility, outline) — it does not materialise any cells. See `rowProperties` for a read-only peek that never creates a record.
- `get rowCount(): number;` — The 1-based index of the last row carrying anything — data or its own formatting — or 0 for an empty sheet. Spans gaps: a value in row 5 makes this 5 even if rows 2–4 are empty. This is the used-range extent, not a populated-row tally (see `actualRowCount`).
- `get actualRowCount(): number;` — The number of rows that hold at least one non-empty cell, ignoring gaps and formatting-only rows.
- `get columnCount(): number;` — The 1-based index of the last column carrying anything — a non-empty cell or its own format properties — or 0 for an empty sheet. The used-range width, mirroring `rowCount` for the other axis: a value in column E makes this 5 even if columns B–D are empty.
- `*columns(): IterableIterator<{
    readonly index: number;
    readonly properties: ColumnProperties;
}>;` — The defined columns in ascending index order, each with its format properties.
- `*rows(): IterableIterator<{
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
}>;` — The rows to serialise, in ascending row order: the union of rows holding cells and rows holding only metadata (a hidden or grouped row need carry no data). Each yields its materialised cells in ascending column order and its format properties, if any. Mirrors how OOXML serialises (`<row>` wrapping `<c>`) and is the writer's row surface.
- `addTable(options: TableOptions): Table;` — Define a table over a range of this sheet. The table's shape invariants (a legal name, at least one column, at least one row) are enforced here; conflicts with the rest of the sheet (e.g. an overlapping merge) are the writer's concern.
- `get tables(): readonly Table[];` — The tables defined on this sheet, in definition order.
- `getTable(name: string): Table | undefined;` — The table with the given name (case-sensitive, the identifier Excel uses), or `undefined`. A table read back from a file is fully hydrated — its rows can be read and appended to.
- `addPivotTable(options: PivotTableOptions): PivotTable;` — Add a pivot table to this (destination) sheet, summarising a source sheet's data. The source is read once, now, so the pivot is a snapshot: later edits to the source do not change it. The supported shape (one summed value field, at least one row and column field) is enforced here.
- `get pivotTables(): readonly PivotTable[];` — The pivot tables hosted on this sheet, in definition order.
- `get loadedPivotTables(): readonly ParsedPivotTable[];` — Pivot tables reconstructed from a loaded package, in the order the reader found them — a read-only inspection view (source range, field roles, value field, aggregation). A pivot authored on this sheet via `addPivotTable` does not appear here; a pivot loaded from a file does not appear in `pivotTables`. The loaded pivots re-emit verbatim through byte-preservation, so this collection is never itself serialised.
- `addCommentThread(thread: CommentThread): void;` — Anchor a threaded conversation to a cell — Excel's modern review comment: an opening message, its replies, and whether the discussion was marked resolved. Distinct from a cell's legacy note (`Cell.note`), and mutually exclusive with one: Excel refuses to put both on one cell, and a cell carrying both is written back as the conversation alone. Every message supplies its own `Comment.id` and `Comment.date`, and names its author by `Comment.personId` into the workbook registry (`Workbook.addPerson`) — the writer has no clock and no id generator, so nothing here is invented and the same workbook always serialises to the same bytes. Every id is normalised to the brace-wrapped upper-case GUID form the format requires, so a `crypto.randomUUID()` is accepted as-is. Message ids must be unique **within this sheet**, because that is the scope in which they mean anything: a reply names its thread by the head's id inside the sheet's own part, and the legacy fallback comment binds its cell by the same id inside the sheet's own comments part. Two sheets reusing one id is therefore harmless and is not rejected — Excel's ids happen to be globally unique, but nothing resolves across a part boundary.
- `get commentThreads(): readonly CommentThread[];` — The threaded conversations on this sheet — Excel's modern review comments (author, timestamp, replies, resolved state, `@mentions`). Empty for a sheet with none. Distinct from a cell's legacy note (`Cell.note`).
- `commentThreadAt(reference: string): CommentThread | undefined;` — The conversation anchored to a cell, or `undefined` when that cell carries none. The reference is canonicalized, so an absolute `"$B$2"` finds the same thread as `"B2"`; it names the *anchor* cell, so a cell merely covered by the anchor's merged region is not a match.
- `addImage(imageId: number, anchor: {
    readonly tl: AnchorPoint;
    readonly br: AnchorPoint;
    readonly editAs?: ImageEditAs;
}): void;` — Anchor a workbook image (the id returned by `Workbook.addImage`) to this sheet. Two shapes: - **Two-cell**: `{tl, br}` spans the rectangle from the top-left grid point to the bottom-right, reflowing as the spanned cells resize. `editAs` (`oneCell` by default) tunes how it follows. - **One-cell**: `{tl, ext}` pins the image at `tl` at a fixed pixel size that the grid never resizes. `ext` is in pixels and converts to EMUs internally. Grid points are 0-based (`{col: 0, row: 0}` is cell A1). A later row/column splice re-pins the anchor to the same logical position.
- `addImageAnchor(imageId: number, anchor: ImageAnchor): void;` — Anchor an image with a pre-built model anchor in the model's own units (EMUs). This is the low-level primitive `addImage` builds on and the reader uses to re-pin an image parsed from a drawing part without a lossy pixel round-trip.
- `removeImage(imageId: number): void;` — Drop every anchor of the given workbook image from this sheet. The image stays registered on the workbook — another sheet may still show it — so only this sheet's anchors are removed; the writer then omits any media no sheet anchors any longer.
- `get images(): readonly AnchoredImage[];` — The images anchored to this sheet, in the order they were added.
- `addBackgroundImage(imageId: number): void;` — Set this sheet's background image to a workbook image (the id `Workbook.addImage` returned). The picture tiles behind the whole grid; it is not anchored to any cell. Passing a new id replaces the previous background.
- `removeBackgroundImage(): void;` — Remove this sheet's background image, if any. The image stays registered on the workbook.
- `get backgroundImageId(): number | undefined;` — The workbook image id set as this sheet's background, or `undefined` when it has none.
- `get preservedReferences(): readonly PreservedWorksheetReference[];` — The worksheet-level references to unmodeled package content preserved for round-tripping.
- `mergeCells(range: string): void;` — Merge a range of cells (`"A1:B2"`). A range that overlaps an already-merged region is rejected — Excel forbids overlapping merges and writes such geometry as a corrupt file. Whole-row/column ranges (`"A:A"`) are unbounded, carry no rectangle, and are not overlap-checked. Any value already sitting in a covered non-anchor cell is discarded, keeping only the top-left anchor's — exactly how Excel collapses a range on merge. Leaving it would emit a populated `<c>` under the `<mergeCell>` ref, the geometry Excel opens with a repair prompt. Covered-cell styles survive (a border spanning the merge is legal), so only the conflicting value is cleared.
- `get merges(): readonly string[];` — The merged ranges on this sheet, in the order they were added.
- `get autoFilter(): AutoFilter | undefined;` — The sheet's autofilter — its range plus any per-column criteria — or `undefined` when the sheet carries none. Setting one turns on the header-row filter dropdowns Excel draws over the range; the writer emits both the sheet's `<autoFilter>` element and the hidden `_FilterDatabase` defined name Excel derives from it. Setting `undefined` clears the filter. A bare range string is the ergonomic common case — `sheet.autoFilter = 'A1:C10'` for dropdowns with no active criteria; pass an `AutoFilter` object to narrow columns. Either way the value is normalised on assignment (range to canonical `A1:C10` form) and the getter returns the structured object. The range must be a bounded rectangle — a whole-row/column reference is not a filterable region and is rejected.
- `unmergeCells(range: string): boolean;` — Remove a merged range previously added with `mergeCells`, returning whether a merge with that exact range string existed. The covering rectangle is dropped alongside it, so a cell the merge had masked addresses independently again. The inverse of `mergeCells`.
- `addDataValidation(sqref: string, rule: DataValidation, options: {
    extended?: boolean;
} = {}): void;` — Attach a data validation to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a space-separated `sqref` of several ranges). The rule is stored once against the range, not copied per covered cell, so a whole-column dropdown stays a single entry. A cell inside the range reports the rule through `dataValidationAt`. Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form (`<x14:dataValidation>`) — the carrier Excel uses for a list source on another sheet and other shapes the standard element cannot express. The reader sets it for a rule found in that form so a round-trip writes it back there instead of silently corrupting the cross-sheet reference.
- `get dataValidations(): readonly DataValidationEntry[];` — The data validations on this sheet, each bound to its target range, in insertion order.
- `addConditionalFormatting(formatting: ConditionalFormatting): void;` — Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref` — one range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one rule set. The block is stored once against the range, defensively copied so the getter never hands back a reference into the caller's object.
- `get conditionalFormattings(): readonly ConditionalFormatting[];` — The conditional formattings on this sheet, each bound to its target range, in insertion order.
- `dataValidationAt(reference: string): DataValidation | undefined;` — The validation covering a cell, or `undefined` when none does. The first added rule whose range contains the cell wins, mirroring how a spreadsheet resolves overlapping validations.
- `spliceRows(start: number, count: number, ...inserts: RowInput[]): void;` — Remove `count` rows starting at the 1-based `start`, then insert the given rows in their place. Rows below the edit shift by `inserts.length - count`: a delete pulls the tail up, an insert pushes it down, and doing both at once is a replace. Each inserted row takes either `RowInput` shape — a positional array from column A, or a key-addressed object — exactly like `addRow`. A `count` larger than the rows present simply clears the tail — it never silently becomes a no-op. Cells carry their full style to the shifted position, and merged ranges shift with the rows they cover.
- `insertRow(pos: number, values: RowInput): void;` — Insert one row of `values` at the 1-based `pos`, shifting the rows at and below it down by one. `values` takes either `RowInput` shape (positional array or keyed object), like `addRow`. Shorthand for `spliceRows``(pos, 0, values)`.
- `addRow(values: RowInput): Cell[];` — Append a row of `values` after the last used row, returning the cells it materialised. The append point is `rowCount`` + 1`, so the row lands below every row that holds data or its own formatting — never overwriting existing content, unlike `insertRow`, which shifts and needs a position. Unlike `spliceRows`, appending shifts nothing, so it never disturbs merges or the rows above. A row takes either shape: a positional array whose values map to columns from A — a hole in a sparse array (`['a', , 'c']`) leaves that column untouched — or a keyed object whose values land under the columns carrying the matching `ColumnProperties.key`.
- `addRows(rows: RowInput[]): Cell[][];` — Append several rows after the last used row in one call, returning the cells materialised for each. The rows stack in order — the first lands at `rowCount`` + 1`, the next directly below it — so a later row never collides with an earlier one even when both are value-less. Each row is an array or a keyed object independently, so a mixed batch is fine. The bulk form of `addRow`.
- `freeze(ySplit = 1, xSplit = 0): void;` — Freeze the top `ySplit` rows and left `xSplit` columns in place; the rest of the sheet scrolls beneath them. `freeze(1)` pins a header row; `freeze(0, 1)` pins the first column. Passing both zero clears the freeze (equivalent to `unfreeze`).
- `unfreeze(): void;` — Clear any frozen split, returning the sheet to a normal (fully scrolling) view.
- `duplicateRow(start: number, options: {
    count?: number;
    insert?: boolean;
} = {}): void;` — Copy the row at the 1-based `start`, `options.count` times (default 1). With `options.insert` (the default) the copies are inserted directly after the source, shifting the rows below — and any merged range there — down by `count`; otherwise the copies overwrite the rows immediately below without shifting. Each copy is a faithful duplicate of the source's values and per-cell styles, and carries no merge of its own, so a range can be merged onto a duplicated row afterwards.
- `spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void;` — Remove `count` columns starting at the 1-based `start`, then insert the given columns in their place — the column analog of `spliceRows`. Columns to the right shift by `inserts.length - count`, keeping their values and styles, and a merged range lying wholly to the right of the edit re-anchors to its new columns. Each inserted column is an array of values indexed by row (index 0 → row 1); an empty array inserts a blank column.
- `insertColumn(pos: number, values: CellValue[]): void;` — Insert one column of `values` at the 1-based `pos`, shifting the columns at and right of it over by one. `values` is an array of values indexed by row (index 0 → row 1), like `addColumn`. Shorthand for `spliceColumns``(pos, 0, values)`.
- `addColumn(values: CellValue[]): Cell[];` — Append a column of `values` after the last used column, returning the cells it materialised. The append point is `columnCount`` + 1`, so the column lands right of every column that holds data or its own formatting — never overwriting existing content, unlike `insertColumn`, which shifts and needs a position. Unlike `spliceColumns`, appending shifts nothing, so it never disturbs merges or the columns to its left. `values` is an array indexed by row (index 0 → row 1); a hole or an explicit `undefined` leaves that row untouched, mirroring `addRow`'s positional-array shape.
- `addColumns(columns: CellValue[][]): Cell[][];` — Append several columns after the last used column in one call, returning the cells materialised for each. The columns stack in order — the first lands at `columnCount`` + 1`, the next directly right of it — so a later column never collides with an earlier one even when both are value-less. The bulk form of `addColumn`.
- `get model(): WorksheetModel;` — A snapshot of this sheet's value and overlay content (see `WorksheetModel`). Reading it and assigning it onto another sheet — `dst.model = src.model` — reproduces the source: merges, cells and their styles, column/row metadata, tables, the autofilter, protection, and the page setup all survive, because the getter emits and the setter consumes exactly the same fields. Identity (`name`, `id`) is not part of the model and is never touched by assignment; nor are attached parts that carry workbook-level identity (images, pivots, byte-preserved charts/drawings) — see `WorksheetModel` for that boundary.
- `protect(password?: string, options: SheetProtectionOptions = {}): void;` — Protect the sheet, making the per-cell `locked`/`hidden` flags enforceable. Without a password the protection is a soft lock any consumer can lift; with one, the password is salted and hashed on the spot (the plaintext is never retained) so lifting the protection requires re-supplying it. `options` names which operations stay available to a user while the sheet is protected; anything unspecified falls to Excel's default for that operation. Re-protecting replaces any prior protection; `unprotect` clears it.
- `unprotect(): void;` — Remove any protection previously set by `protect`.
- `get protection(): SheetProtection | undefined;` — The sheet's protection, or `undefined` if the sheet is unprotected.

---

### `WorksheetModel`

<sub>interface</sub>

A serialisable snapshot of a worksheet's value and overlay content — its cells and their styles,
the column/row/page metadata, and the sheet-level overlays (merges, data validations, conditional
formattings, tables, the autofilter, protection). `Worksheet.model` exports one; assigning
it back reproduces that content. The getter and setter cover exactly the same fields, so a
`dst.model = src.model` round-trip drops none of it — an export field the import ignored would
silently lose data, the historical merge-loss failure this contract exists to prevent. Both
directions are driven from one field table (`core/worksheet-model.ts`), which the compiler proves
covers every field below, so adding a field here without wiring it fails the build.

Out of scope by design: content that carries workbook-level identity rather than pure sheet
state — anchored and background images (their bytes live on the `Workbook`), pivot tables
(their source references a live worksheet), and byte-preserved parts (charts, vector drawings,
slicers) kept verbatim for round-tripping. These stay with their source sheet; a model assignment
neither copies nor clears them.

```ts
interface WorksheetModel {
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
    columns: {
        index: number;
        properties: ColumnProperties;
    }[];
    rows: {
        number: number;
        properties: RowProperties;
    }[];
    cells: CellModel[];
    merges: string[];
    dataValidations: DataValidationEntry[];
    conditionalFormattings: ConditionalFormatting[];
    tables: TableOptions[];
    autoFilter: AutoFilter | undefined;
    protection: SheetProtection | undefined;
}
```

---

### `WorksheetProperties`

<sub>interface</sub>

Format defaults applied to every row/column that carries no explicit override.

```ts
interface WorksheetProperties {
    defaultRowHeight?: number;
    defaultColWidth?: number;
}
```

---

### `WorksheetState`

<sub>interface</sub>

```ts
interface WorksheetState {
    readonly state: 'visible' | 'hidden' | 'veryHidden';
}
```
