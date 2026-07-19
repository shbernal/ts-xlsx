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
  get fill(): Fill | undefined;
  set fill(fill: Fill | undefined);
  get numFmt(): string | undefined;
  set numFmt(numFmt: string | undefined);
  get font(): Partial<Font> | undefined;
  set font(font: Partial<Font> | undefined);
  get border(): Border | undefined;
  set border(border: Border | undefined);
  get alignment(): Alignment | undefined;
  set alignment(alignment: Alignment | undefined);
  get protection(): Protection | undefined;
  set protection(protection: Protection | undefined);
  get quotePrefix(): boolean | undefined;
  set quotePrefix(quotePrefix: boolean | undefined);
  get namedStyleId(): number | undefined;
  set namedStyleId(namedStyleId: number | undefined);
  get note(): string | undefined;
  set note(note: string | undefined);
}
```

**Members**

- `readonly row: number;` — 1-based row index.
- `readonly col: number;` — 1-based column index.
- `get address(): string;` — Canonical A1 address of this cell (`"B3"`).
- `get value(): CellValue;` — The cell's value; `null` when empty. Assigning `undefined` clears it.
- `get type(): ValueType;` — The observable `ValueType` of the current value.
- `get fill(): Fill | undefined;` — The cell's background fill, or `undefined` when it has none.
- `get numFmt(): string | undefined;` — The cell's number-format code (`"0.00%"`, a custom accounting format, …), or `undefined` for the General format. Stored verbatim: the invariant form Excel persists — `.` decimal, `,` grouping, `/` date separator — is neither localized nor rewritten, so the code round-trips character-for-character. A cell that also carries a column-level format keeps both, so overriding one facet never drops the other.
- `get font(): Partial<Font> | undefined;` — The cell's font — bold/italic/underline, size, colour, typeface — as a partial set of the facets that differ from the default (only the facets actually set are carried, exactly as OOXML stores them). `undefined` means the cell uses the workbook default font.
- `get border(): Border | undefined;` — The cell's border — the line style and colour of each side — or `undefined` when the cell has none. An absent edge within a border means that side is unbordered, so reading a cell never fabricates a border it does not have.
- `get alignment(): Alignment | undefined;` — The cell's alignment — how its content sits within the cell, plus the wrap/shrink flags — or `undefined` when it uses the defaults. The boolean flags are off unless explicitly set, so a cell that never enabled wrapping never reads back wrapped.
- `get protection(): Protection | undefined;` — The cell's protection — its locked/hidden flags, enforced only once the sheet is protected — or `undefined` when the cell carries neither. `locked` defaults to on in OOXML, so a cell that never touched protection is implicitly locked and reads back as `undefined`, not as `{locked: true}`; the flag only becomes explicit when a cell is unlocked.
- `get quotePrefix(): boolean | undefined;` — The quote-prefix flag: when set, a spreadsheet stores the cell's content as literal text even when it looks like a formula or number, and shows a leading apostrophe in the formula bar without that apostrophe being part of the stored value. `undefined` (or `false`) when unset. It is a cell-format flag — an attribute on the cell's `xf` record — so it composes independently of the value.
- `get namedStyleId(): number | undefined;` — The index of the `Workbook.namedStyles named cell style` this cell links to (its OOXML `xfId`), or `undefined` when the cell references no named style beyond the default. The cell inherits any facet its own direct format leaves unset from that named style; the reader resolves the effective look onto the cell's own facets, and this link is preserved so a round-trip keeps the cell tied to its named style rather than flattening it away.
- `get note(): string | undefined;` — The cell's note (comment) as plain text, or `undefined` when it carries none. A note is metadata anchored to the cell, independent of its value: a cell can hold a note while empty, and clearing the value leaves the note intact. A structural edit that shifts the cell carries the note along to its new position.
