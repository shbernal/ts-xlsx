# Addresses & ranges

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellAddress`

<sub>interface</sub>

A decoded single-cell reference. An axis the reference omits is `undefined`.

```ts
interface CellAddress {
  /** Canonical A1 form with `$` anchors stripped — e.g. `"B2"`, `"1"`, `"A"`. */
  readonly address: string;
  /** 1-based column, or `undefined` for a row-only reference (`$1`). */
  readonly col: number | undefined;
  /** 1-based row, or `undefined` for a column-only reference (`$A`). */
  readonly row: number | undefined;
}
```

---

### `columnToNumber`

<sub>function</sub>

Convert column letters to a 1-based number (`"A" → 1`, `"AA" → 27`).

```ts
function columnToNumber(letters: string): number;
```

---

### `decodeAddress`

<sub>function</sub>

Decode a single cell/row/column reference into `{address, col, row}`. Anchoring
`$` signs are accepted and dropped; an absent axis is `undefined`.

```ts
function decodeAddress(reference: string): CellAddress;
```

**Throws** — if the reference mentions neither a column nor a row.

---

### `decodeRange`

<sub>function</sub>

Decode a range reference (`A1:B2`, `$1:$1`, `Sheet1!$A:$A`) into its corners and
canonical dimensions. A single reference collapses to a degenerate range whose
corners coincide.

```ts
function decodeRange(reference: string): RangeAddress;
```

---

### `encodeAddress`

<sub>function</sub>

Encode a 1-based `col`/`row` pair into its canonical A1 address (`"B2"`).

```ts
function encodeAddress(col: number, row: number): string;
```

---

### `MAX_COLUMN`

<sub>const</sub>

Excel's column bounds: `A` (1) through `XFD` (16384).

```ts
const MAX_COLUMN: 16384
```

---

### `MAX_ROW`

<sub>const</sub>

Excel's row bound: 1 through 1048576. The other axis of `MAX_COLUMN`.

```ts
const MAX_ROW: 1048576
```

---

### `numberToColumn`

<sub>function</sub>

Convert a 1-based column number to its letters (`1 → "A"`, `27 → "AA"`).

```ts
function numberToColumn(n: number): string;
```

---

### `RangeAddress`

<sub>interface</sub>

A decoded range reference. Corners are the min/max of the endpoints per axis;
an axis neither endpoint mentions (a whole-row or whole-column range) is
`undefined` on every corner and simply absent from `dimensions`.

```ts
interface RangeAddress {
  readonly top: number | undefined;
  readonly left: number | undefined;
  readonly bottom: number | undefined;
  readonly right: number | undefined;
  /** The originating sheet, present only when the reference carried one. */
  readonly sheetName?: string;
  readonly tl: CellAddress;
  readonly br: CellAddress;
  /** Canonical `tl:br` form — `"A1:B2"`, `"1:1"` (rows), `"A:A"` (columns). */
  readonly dimensions: string;
}
```
