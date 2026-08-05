# Cell

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Cell`

<sub>class</sub>

A single cell owns its value and every style facet outright. Each facet below — fill, number format,
font, border, alignment, protection, quote-prefix, and note — is held in the cell's own field and
*replaced* (never mutated in place) by its setter, so a facet set on one cell never aliases or bleeds
onto its row, column, or sheet siblings. Each facet's own doc covers only what is specific to it.

```ts
class Cell {
  readonly row: number;
  readonly col: number;
  get address(): string;
  get value(): CellValue;
  set value(value: CellValue | undefined);
  get type(): ValueType;
  setRichText(runs: readonly RichTextRun[]): void;
  get style(): CellStyle;
  set style(style: Readonly<CellStyle>);
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
  get quotePrefix(): boolean | undefined;
  set quotePrefix(quotePrefix: boolean | undefined);
  get note(): string | undefined;
  set note(note: string | undefined);
}
```

**Members**

#### `Cell.row`

```ts
readonly row: number;
```

1-based row index.

#### `Cell.col`

```ts
readonly col: number;
```

1-based column index.

#### `Cell.address`

```ts
get address(): string;
```

Canonical A1 address of this cell (`"B3"`).

#### `Cell.value`

```ts
get value(): CellValue;
set value(value: CellValue | undefined);
```

The cell's value; `null` when empty. Assigning `undefined` clears it.

#### `Cell.type`

```ts
get type(): ValueType;
```

The observable `ValueType` of the current value.

#### `Cell.setRichText`

```ts
setRichText(runs: readonly RichTextRun[]): void;
```

Assign rich text whose runs **inherit this cell's font**, so a run needs to state only what it
changes: `setRichText([{text: 'Note:', font: {bold: true}}, {text: ' the rest'}])` keeps the
cell's typeface, size and colour throughout and bolds the first run.

This exists because a run's format element does **not** inherit anything. A `<rPr>` is a
*complete* character format, and any facet it omits falls back to the workbook default font
(`Workbook.defaultFont`) — not to the cell's. Verified against Excel: a cell set to
Courier New 16 whose first run carries only `<b/>` renders that run in the workbook default face
at the default size, bold, while the rest of the cell renders Courier New 16. So a run authored
as `{bold: true}` beside a styled cell silently loses the face, which is the format's rule rather
than a bug — and the reason this is a helper rather than a change to how runs are written.

Composition is per facet: a facet the run names wins, one it omits comes from the cell. Assigning
`value` directly stays the bare path, for a caller who wants a run that deliberately falls back
to the workbook default.

A cell that names no font of its own needs no composition — an omitted facet already falls back
to the workbook default, which is exactly what such a cell renders in — so the runs pass through
unchanged.

#### `Cell.style`

```ts
get style(): CellStyle;
set style(style: Readonly<CellStyle>);
```

The cell's full style — fill, number format, font, border, alignment, and protection — as one
`CellStyle`, for restyling a cell wholesale without importing `applyCellStyle`
separately (mirrors `Worksheet.model`'s getter/setter pair for the whole sheet). The
getter carries only the facets this cell has set (the same shape `cellToModel` emits);
the setter lays each facet `style` carries onto this cell — like every per-facet setter, it
replaces that facet outright but leaves a facet `style` omits untouched, so `cell.style = {...}`
composes with prior per-facet sets rather than clearing them wholesale.

#### `Cell.fill`

```ts
get fill(): Fill | undefined;
set fill(fill: Fill | undefined);
```

The cell's background fill, or `undefined` when it has none.

#### `Cell.numFmt`

```ts
get numFmt(): string | undefined;
set numFmt(numFmt: string | undefined);
```

The cell's number-format code (`"0.00%"`, a custom accounting format, …), or
`undefined` for the General format. Stored verbatim: the invariant form Excel
persists — `.` decimal, `,` grouping, `/` date separator — is neither localized
nor rewritten, so the code round-trips character-for-character. A cell that also carries
a column-level format keeps both, so overriding one facet never drops the other.

#### `Cell.font`

```ts
get font(): Font | undefined;
set font(font: Font | undefined);
```

The cell's font — bold/italic/underline, size, colour, typeface — as a partial set
of the facets that differ from the default (only the facets actually set are carried,
exactly as OOXML stores them). `undefined` means the cell uses the workbook default font.

#### `Cell.border`

```ts
get border(): Border | undefined;
set border(border: Border | undefined);
```

The cell's border — the line style and colour of each side — or `undefined` when the
cell has none. An absent edge within a border means that side is unbordered, so reading
a cell never fabricates a border it does not have.

#### `Cell.alignment`

```ts
get alignment(): Alignment | undefined;
set alignment(alignment: Alignment | undefined);
```

The cell's alignment — how its content sits within the cell, plus the wrap/shrink flags —
or `undefined` when it uses the defaults. The boolean flags are off unless explicitly set,
so a cell that never enabled wrapping never reads back wrapped.

#### `Cell.protection`

```ts
get protection(): Protection | undefined;
set protection(protection: Protection | undefined);
```

The cell's protection — its locked/hidden flags, enforced only once the sheet is protected —
or `undefined` when the cell carries neither. `locked` defaults to on in OOXML, so a cell
that never touched protection is implicitly locked and reads back as `undefined`, not as
`{locked: true}`; the flag only becomes explicit when a cell is unlocked.

#### `Cell.quotePrefix`

```ts
get quotePrefix(): boolean | undefined;
set quotePrefix(quotePrefix: boolean | undefined);
```

The quote-prefix flag: when set, a spreadsheet stores the cell's content as literal text even
when it looks like a formula or number, and shows a leading apostrophe in the formula bar without
that apostrophe being part of the stored value. `undefined` (or `false`) when unset. It is a
cell-format flag — an attribute on the cell's `xf` record — so it composes independently of the
value.

#### `Cell.note`

```ts
get note(): string | undefined;
set note(note: string | undefined);
```

The cell's note (comment) as plain text, or `undefined` when it carries none. A note is
metadata anchored to the cell, independent of its value: a cell can hold a note while empty,
and clearing the value leaves the note intact. A structural edit that shifts the cell carries the
note along to its new position.
