# Images

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `AnchoredImage`

<sub>interface</sub>

An image pinned to a worksheet: which workbook media it shows (`imageId`) and where.

```ts
interface AnchoredImage {
    readonly imageId: number;
    readonly anchor: ImageAnchor;
}
```

---

### `AnchorPoint`

<sub>interface</sub>

A point in the drawing grid: a 0-based column and row, plus an EMU offset into that cell.
The offsets default to zero, pinning the point to the cell's top-left corner.

```ts
interface AnchorPoint {
    readonly col: number;
    readonly row: number;
    readonly colOff?: number;
    readonly rowOff?: number;
}
```

---

### `Extent`

<sub>interface</sub>

A fixed image size in EMUs — the extent of a one-cell anchor, which pixel dimensions convert into
via `PX_TO_EMU`.

```ts
interface Extent {
    readonly cx: number;
    readonly cy: number;
}
```

---

### `ImageAnchor`

<sub>type</sub>

Where an image sits on the grid: a rectangle between two cells, or a point plus a fixed extent.

```ts
type ImageAnchor = TwoCellAnchor | OneCellAnchor;
```

---

### `ImageEditAs`

<sub>type</sub>

How a two-cell-anchored image tracks edits to the cells it spans. `twoCell` moves and resizes with
them; `oneCell` moves but keeps its size; `absolute` is pinned to the page and does neither. Excel
defaults to `oneCell` when the attribute is omitted.

```ts
type ImageEditAs = 'oneCell' | 'twoCell' | 'absolute';
```

---

### `isOneCellAnchor`

<sub>function</sub>

Narrow an anchor to its one-cell (fixed-extent) form; the complement is `TwoCellAnchor`.

```ts
function isOneCellAnchor(anchor: ImageAnchor): anchor is OneCellAnchor;
```

---

### `OneCellAnchor`

<sub>interface</sub>

A one-cell anchor: a single top-left grid point plus a fixed extent. The image keeps its size as
the grid resizes, moving only with its anchor cell. `editAs` is a two-cell-only attribute and has
no place here.

```ts
interface OneCellAnchor {
    readonly from: AnchorPoint;
    readonly ext: Extent;
    readonly rotation?: number;
}
```

---

### `PX_TO_EMU`

<sub>const</sub>

EMUs per pixel at Excel's notional 96 DPI (914400 EMU/inch ÷ 96 px/inch). The conversion is
DPI-independent by construction: a pixel extent is a fixed physical size regardless of screen.

```ts
const PX_TO_EMU: 9525
```

---

### `TwoCellAnchor`

<sub>interface</sub>

A two-cell anchor: the image's top-left (`from`) and bottom-right (`to`) grid points. The image
fills the rectangle between them and reflows as the intervening rows/columns resize; `editAs`
selects how strictly it follows.

```ts
interface TwoCellAnchor {
    readonly from: AnchorPoint;
    readonly to: AnchorPoint;
    readonly editAs?: ImageEditAs;
    readonly rotation?: number;
}
```

---

### `WorkbookImage`

<sub>interface</sub>

A picture's bytes and its file kind, as held in the workbook's media registry.

```ts
interface WorkbookImage {
    readonly extension: string;
    readonly data: Uint8Array;
}
```
