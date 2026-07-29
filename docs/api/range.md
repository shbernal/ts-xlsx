# Range

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Range`

<sub>class</sub>

A rectangular block of a worksheet's cells: `sheet.getRange('B2:D5')`.

Cheap and stateless — constructing one creates no cells and does not extend the used range.
`addresses` walks the block without materialising anything; `cells` reports only what
already exists.

Bounds are **inclusive first/last**, never start-and-count. That is the convention every
range-shaped accessor in this library follows, so the three axes cannot disagree about what a
pair of numbers means.

```ts
class Range {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
  get sheet(): Worksheet;
  get address(): string;
  get rowCount(): number;
  get columnCount(): number;
  get cellCount(): number;
  contains(row: number, col: number): boolean;
  *addresses(): IterableIterator<string>;
  get cells(): readonly Cell[];
}
```

**Members**

- `readonly top: number;` — 1-based row of the block's top edge. Fixed for this handle's lifetime.
- `readonly left: number;` — 1-based column of the block's left edge. Fixed for this handle's lifetime.
- `readonly bottom: number;` — 1-based row of the block's bottom edge, inclusive.
- `readonly right: number;` — 1-based column of the block's right edge, inclusive.
- `get sheet(): Worksheet;` — The worksheet this block belongs to.
- `get address(): string;` — Canonical `tl:br` A1 form — `"B2:D5"`. A one-cell block still reads as `"B2:B2"`.
- `get rowCount(): number;` — Rows spanned, inclusive of both edges.
- `get columnCount(): number;` — Columns spanned, inclusive of both edges.
- `get cellCount(): number;` — Cells the block covers — `rowCount * columnCount`, whether or not they exist yet.
- `contains(row: number, col: number): boolean;` — Whether a 1-based position falls inside the block.
- `*addresses(): IterableIterator<string>;` — Every address the block covers, row-major (`B2`, `C2`, `D2`, `B3`, …). Materialises nothing, so this is the cheap way to walk a large block — and, being a generator, it can be abandoned part-way without having built the whole list.
- `get cells(): readonly Cell[];` — The block's **materialised** cells, row-major. Sparse: a position nothing has ever written to is simply absent, which is what distinguishes "never written" from a cell holding `null`. Reading this creates nothing — mirroring `Column.cells`.
