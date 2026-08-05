# Range

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Range`

<sub>class</sub>

A rectangular block of a worksheet's cells: `sheet.getRange('B2:D5')`.

Cheap and stateless — constructing one creates no cells and does not extend the used range.
[`addresses`](./range.md#rangeaddresses) walks the block without materialising anything; [`cells`](./range.md#rangecells) reports only what
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
  get style(): CellStyle;
  set style(style: Readonly<CellStyle>);
  clearStyle(): void;
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
}
```

**Members**

#### `Range.top`

```ts
readonly top: number;
```

1-based row of the block's top edge. Fixed for this handle's lifetime.

#### `Range.left`

```ts
readonly left: number;
```

1-based column of the block's left edge. Fixed for this handle's lifetime.

#### `Range.bottom`

```ts
readonly bottom: number;
```

1-based row of the block's bottom edge, inclusive.

#### `Range.right`

```ts
readonly right: number;
```

1-based column of the block's right edge, inclusive.

#### `Range.sheet`

```ts
get sheet(): Worksheet;
```

The worksheet this block belongs to.

#### `Range.address`

```ts
get address(): string;
```

Canonical `tl:br` A1 form — `"B2:D5"`. A one-cell block still reads as `"B2:B2"`.

#### `Range.rowCount`

```ts
get rowCount(): number;
```

Rows spanned, inclusive of both edges.

#### `Range.columnCount`

```ts
get columnCount(): number;
```

Columns spanned, inclusive of both edges.

#### `Range.cellCount`

```ts
get cellCount(): number;
```

Cells the block covers — `rowCount * columnCount`, whether or not they exist yet.

#### `Range.contains`

```ts
contains(row: number, col: number): boolean;
```

Whether a 1-based position falls inside the block.

#### `Range.addresses`

```ts
*addresses(): IterableIterator<string>;
```

Every address the block covers, row-major (`B2`, `C2`, `D2`, `B3`, …). Materialises nothing, so
this is the cheap way to walk a large block — and, being a generator, it can be abandoned
part-way without having built the whole list.

#### `Range.cells`

```ts
get cells(): readonly Cell[];
```

The block's **materialised** cells, row-major. Sparse: a position nothing has ever written to is
simply absent, which is what distinguishes "never written" from a cell holding `null`. Reading
this creates nothing — mirroring [`Column.cells`](./column.md#columncells).

#### `Range.style`

```ts
get style(): CellStyle;
set style(style: Readonly<CellStyle>);
```

The block's style, facet by facet — the counterpart of [`Cell.style`](./cell.md#cellstyle) over a rectangle, with
the same semantics in both directions.

**Reading** reports a facet only when *every* position in the block carries a structurally
identical one, and `undefined` when they differ or any position is still empty. A block styled
through this handle therefore reads back what was written; a block whose cells disagree says so
rather than picking a corner's answer and passing it off as the whole.

**Writing** lays each facet the payload names onto every cell, leaving facets it omits untouched
— exactly what `cell.style = {...}` does, so this composes with prior styling instead of clearing
it. Use [`clearStyle`](./range.md#rangeclearstyle) first for a wholesale replace.

Writing **materialises** every position in the block, because a styled-but-valueless cell is the
only way an empty cell renders with a fill: skipping the holes would leave gaps in a header band.
The cost is bounded by construction — a range is always a bounded rectangle, and whole-axis
styling belongs to [`Worksheet.getColumn`](./worksheet.md#worksheetgetcolumn)/[`Worksheet.getRow`](./worksheet.md#worksheetgetrow) instead. [`cellCount`](./range.md#rangecellcount)
is the exact number of cells a write will create.

#### `Range.clearStyle`

```ts
clearStyle(): void;
```

Strip every style facet from every cell in the block, leaving values untouched. Assigning
[`style`](./range.md#rangestyle) composes, so this is how a wholesale replace is said: `clearStyle()` then assign.
Materialises nothing — a cell that does not exist carries no style to clear.

#### `Range.fill`

```ts
get fill(): Fill | undefined;
set fill(fill: Fill | undefined);
```

Fill applied to every cell in the block; `undefined` when they do not all agree.

#### `Range.numFmt`

```ts
get numFmt(): string | undefined;
set numFmt(numFmt: string | undefined);
```

Number format applied to every cell in the block; `undefined` when they do not all agree.

#### `Range.font`

```ts
get font(): Font | undefined;
set font(font: Font | undefined);
```

Font applied to every cell in the block; `undefined` when they do not all agree.

#### `Range.border`

```ts
get border(): Border | undefined;
set border(border: Border | undefined);
```

Border applied to every cell in the block; `undefined` when they do not all agree.

#### `Range.alignment`

```ts
get alignment(): Alignment | undefined;
set alignment(alignment: Alignment | undefined);
```

Alignment applied to every cell in the block; `undefined` when they do not all agree.

#### `Range.protection`

```ts
get protection(): Protection | undefined;
set protection(protection: Protection | undefined);
```

Protection flags applied to every cell in the block; `undefined` when they do not all agree.
