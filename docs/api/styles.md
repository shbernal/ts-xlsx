# Styles

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Alignment`

<sub>interface</sub>

A cell's alignment. Every facet is optional and independent; an absent facet means the cell
takes Excel's default for it. The boolean flags default to off — a cell that never enabled
`wrapText`/`shrinkToFit` must never read back with them on. `textRotation` is in degrees
(0–180, where 91–180 encodes -1° to -90°); `indent` is a non-negative indent level.

```ts
interface Alignment {
    readonly horizontal?: HorizontalAlignment;
    readonly vertical?: VerticalAlignment;
    readonly textRotation?: number;
    readonly wrapText?: boolean;
    readonly indent?: number;
    readonly shrinkToFit?: boolean;
    readonly readingOrder?: number;
}
```

---

### `Border`

<sub>interface</sub>

A cell's border. Each of the four sides plus the diagonal is an independent edge; an
absent edge means that side has no border (it is not rendered), so a cell bordered on
one side never implies the other three. `diagonalUp`/`diagonalDown` select which way a
present diagonal edge runs.

```ts
interface Border {
    readonly left?: BorderEdge;
    readonly right?: BorderEdge;
    readonly top?: BorderEdge;
    readonly bottom?: BorderEdge;
    readonly diagonal?: BorderEdge;
    readonly diagonalUp?: boolean;
    readonly diagonalDown?: boolean;
}
```

---

### `BorderEdge`

<sub>interface</sub>

One edge of a cell border: its line style, and optionally the line colour.

```ts
interface BorderEdge {
    readonly style: BorderStyle;
    readonly color?: Color;
}
```

---

### `BorderStyle`

<sub>type</sub>

Line styles a cell border edge can take, as OOXML's `ST_BorderStyle` enumerates them.
`none` is the absence of an edge and is expressed by omitting the edge, not by this value.

```ts
type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
```

---

### `CellStyle`

<sub>interface</sub>

The six direct-format facets a cell can carry — its fill, number format, font, border, alignment,
and protection. Every facet is optional and independent: a cell sets only the facets it overrides
and inherits the rest. This one tuple is the unit of style throughout the library, so the
interfaces that carry a cell's formatting compose it rather than re-listing the fields — a column,
table column, or named style whose facets *default* the cells that leave them unset (see
`ColumnProperties`, `NamedCellStyle`), and a cell's own resolved format. Because they
share this type, "add a facet" is a single edit here and the compiler enforces that no read/write
path silently drops one — the round-trip symmetry the merge-loss contract depends on.

```ts
interface CellStyle {
    fill?: Fill | undefined;
    numFmt?: string | undefined;
    font?: Partial<Font> | undefined;
    border?: Border | undefined;
    alignment?: Alignment | undefined;
    protection?: Protection | undefined;
}
```

---

### `Color`

<sub>interface</sub>

A colour, expressed as an ARGB hex string (`"FF0000FF"`) or an indexed theme colour.

```ts
interface Color {
    readonly argb?: string;
    readonly theme?: number;
    readonly tint?: number;
    readonly indexed?: number;
}
```

---

### `Fill`

<sub>type</sub>

A cell/row background fill: a flat pattern or a colour gradient.

```ts
type Fill = PatternFill | GradientFill;
```

---

### `FillPatternType`

<sub>type</sub>

Fill pattern kinds, as OOXML's `ST_PatternType` enumerates them. `none` is the
absence of a fill; `solid` paints the whole cell with the foreground colour (the
common case). The remaining hatch patterns are carried for fidelity on read.

```ts
type FillPatternType = 'none' | 'solid' | 'gray125' | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray0625' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis';
```

---

### `Font`

<sub>interface</sub>

A font, as it applies to a cell or a single rich-text run.

```ts
interface Font {
    readonly name: string;
    readonly size: number;
    readonly family: number;
    readonly scheme: 'minor' | 'major' | 'none';
    readonly charset: number;
    readonly color: Color;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly underline: UnderlineStyle;
    readonly strike: boolean;
    readonly outline: boolean;
    readonly vertAlign: FontVerticalAlignment;
}
```

---

### `FontVerticalAlignment`

<sub>type</sub>

Vertical alignment of a font relative to the baseline (super/subscript).

```ts
type FontVerticalAlignment = 'superscript' | 'subscript';
```

---

### `GradientFill`

<sub>interface</sub>

A gradient fill, as OOXML's `CT_GradientFill`. A `linear` gradient runs at `degree`
degrees across the cell; a `path` gradient radiates from an inner rectangle whose
insets are `left`/`right`/`top`/`bottom` (each a fraction in `[0, 1]`). The `stops`
place colours along the axis; a well-formed gradient names at least two.

```ts
interface GradientFill {
    readonly type: 'gradient';
    readonly gradient: 'linear' | 'path';
    readonly degree?: number;
    readonly left?: number;
    readonly right?: number;
    readonly top?: number;
    readonly bottom?: number;
    readonly stops: readonly GradientStop[];
}
```

---

### `GradientStop`

<sub>interface</sub>

A colour stop in a gradient, at a fractional `position` in `[0, 1]` along the gradient axis.

```ts
interface GradientStop {
    readonly position: number;
    readonly color: Color;
}
```

---

### `HorizontalAlignment`

<sub>type</sub>

How a cell's content sits horizontally within its bounds, as OOXML's `ST_HorizontalAlignment`
enumerates it. `general` is the type-dependent default (text left, numbers right) and reads
back as no explicit horizontal alignment.

```ts
type HorizontalAlignment = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
```

---

### `PatternFill`

<sub>interface</sub>

A pattern fill. For a `solid` fill the visible colour is the pattern *foreground*
(`fgColor`) — OOXML's counter-intuitive rule — while `bgColor` is the automatic
indexed placeholder.

```ts
interface PatternFill {
    readonly type: 'pattern';
    readonly pattern: FillPatternType;
    readonly fgColor?: Color;
    readonly bgColor?: Color;
}
```

---

### `Protection`

<sub>interface</sub>

A cell's protection state, enforced only when the worksheet itself is protected — the flags
do nothing on an unprotected sheet. `locked` defaults to TRUE in OOXML (every cell is locked
unless told otherwise), so the meaningful, information-carrying state is an explicitly
*unlocked* cell (`locked: false`); marking a cell locked merely restates the default and
records nothing. `hidden` (defaults false) hides the cell's formula from the formula bar of a
protected sheet. A cell that never set either flag reads back with neither.

```ts
interface Protection {
    readonly locked?: boolean;
    readonly hidden?: boolean;
}
```

---

### `UnderlineStyle`

<sub>type</sub>

Underline can be a plain flag or one of Excel's named underline styles.

```ts
type UnderlineStyle = boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
```

---

### `VerticalAlignment`

<sub>type</sub>

How a cell's content sits vertically within its bounds, as OOXML's `ST_VerticalAlignment`
enumerates it.

```ts
type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
```
