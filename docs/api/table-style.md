# Table Style

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `isTableStyleElementType`

<sub>function</sub>

```ts
function isTableStyleElementType(value: string): value is TableStyleElementType;
```

---

### `STRIPE_ELEMENT_TYPES`

<sub>const</sub>

The four element types banded across several rows or columns — the only ones `TableStyleElement.size` means anything on.

```ts
const STRIPE_ELEMENT_TYPES: ReadonlySet<"wholeTable" | "headerRow" | "totalRow" | "firstColumn" | "lastColumn" | "firstRowStripe" | "secondRowStripe" | "firstColumnStripe" | "secondColumnStripe" | "firstHeaderCell" | "lastHeaderCell" | "firstTotalCell" | "lastTotalCell" | "firstSubtotalColumn" | "secondSubtotalColumn" | "thirdSubtotalColumn" | "firstSubtotalRow" | "secondSubtotalRow" | "thirdSubtotalRow" | "blankRow" | "firstColumnSubheading" | "secondColumnSubheading" | "thirdColumnSubheading" | "firstRowSubheading" | "secondRowSubheading" | "thirdRowSubheading" | "pageFieldLabels" | "pageFieldValues">
```

---

### `TABLE_STYLE_ELEMENT_TYPES`

<sub>const</sub>

The regions a table style can format (`ST_TableStyleType`).

The first thirteen apply to a **table**; the rest style a **pivot table**, which has regions a
table does not have (subtotal rows, page-field labels, subheadings). Both live in the same
enumeration and the same `<tableStyle>` element — what decides which regions a consumer honours is
the style's own `table`/`pivot` flags, not the element names — so the type carries all of them
rather than splitting into two enumerations that a caller would have to choose between up front.

```ts
const TABLE_STYLE_ELEMENT_TYPES: readonly ["wholeTable", "headerRow", "totalRow", "firstColumn", "lastColumn", "firstRowStripe", "secondRowStripe", "firstColumnStripe", "secondColumnStripe", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell", "firstSubtotalColumn", "secondSubtotalColumn", "thirdSubtotalColumn", "firstSubtotalRow", "secondSubtotalRow", "thirdSubtotalRow", "blankRow", "firstColumnSubheading", "secondColumnSubheading", "thirdColumnSubheading", "firstRowSubheading", "secondRowSubheading", "thirdRowSubheading", "pageFieldLabels", "pageFieldValues"]
```

---

### `TableStyle`

<sub>interface</sub>

A custom table style, ready to be registered on a workbook and named by a table's
`TableStyleInfo.name`.

Elements are applied in the order ECMA-376 fixes, not the order they are written here: whole table,
then the column stripes, then the row stripes, then last/first column, header row, total row, and
the four corner cells. So a row stripe wins over a column stripe, and both win over the whole-table
formatting — worth knowing when a stripe colour appears not to take.

```ts
interface TableStyle {
  /** The name a table references, and the name Excel shows in its style gallery. */
  readonly name: string;
  /** The regions this style formats. An element left out is not styled by it. */
  readonly elements: Readonly<Partial<Record<TableStyleElementType, TableStyleElement>>>;
  /** Whether the style is offered for tables. Defaults to true. */
  readonly table?: boolean | undefined;
  /** Whether the style is offered for pivot tables. Defaults to true. */
  readonly pivot?: boolean | undefined;
}
```

---

### `TableStyleElement`

<sub>interface</sub>

How one region of a table is formatted: a `DifferentialStyle` laid over whatever the cells
already carry, plus — for a stripe — how many rows or columns wide one band is.

A `numFmt` here is carried faithfully but has no visible effect: Excel's own table-style element
exposes a font, an interior and borders, and nothing for a number format. See
`DifferentialStyle`.

```ts
interface TableStyleElement extends DifferentialStyle {
  /**
   * The band width, in rows or columns, for a striped element — `2` makes each band two rows deep.
   * Defaults to 1.
   *
   * Meaningful **only** on the four stripe types ({@link STRIPE_ELEMENT_TYPES}); ECMA-376 says so and
   * Excel ignores it elsewhere. Setting it on any other element is rejected rather than silently
   * dropped: a caller who wrote it meant something by it, and a value that vanishes into a file that
   * still opens cleanly is the kind of bug nobody finds.
   */
  readonly size?: number | undefined;
}
```

---

### `TableStyleElementType`

<sub>type</sub>

One region of a table or pivot that a table style can format.

```ts
type TableStyleElementType = (typeof TABLE_STYLE_ELEMENT_TYPES)[number];
```
