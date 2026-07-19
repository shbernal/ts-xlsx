# Page Setup

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `HeaderFooter`

<sub>interface</sub>

Page header/footer text, one string per page class. Excel only honours the even- and
first-page variants when the writer also sets the gating flags (`differentOddEven`,
`differentFirst`); the writer derives those from which variants are present. An empty
object means the element is omitted entirely.

```ts
interface HeaderFooter {
    oddHeader?: string;
    oddFooter?: string;
    evenHeader?: string;
    evenFooter?: string;
    firstHeader?: string;
    firstFooter?: string;
}
```

---

### `PageBreak`

<sub>interface</sub>

A manual page break (`<brk>`). For a row break, `id` is the row the layout splits *before*; for a
column break it is the column. `max` bounds the break's extent across the other axis (Excel writes
the last row/column index) and `man` marks it author-set rather than automatic — the model preserves
whatever the source carried so a round-trip reproduces the break's span exactly.

```ts
interface PageBreak {
    readonly id: number;
    readonly max?: number;
    readonly man?: boolean;
}
```

---

### `PageMargins`

<sub>interface</sub>

Print margins, in inches. OOXML's `<pageMargins>` requires all six to be present, but
the model stores only what the caller set; the writer fills the untouched ones with
valid defaults. An empty object means the element is omitted entirely.

```ts
interface PageMargins {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
}
```

---

### `PageSetup`

<sub>interface</sub>

Print-scaling and orientation settings. These map onto two OOXML elements: `fitToPage` is the
`<pageSetUpPr>` flag (a `<sheetPr>` child) that switches Excel from fixed-zoom to fit-to-page
scaling, while the rest are `<pageSetup>` attributes. Excel honours `scale` only when `fitToPage`
is off and the `fitToWidth`/`fitToHeight` page counts only when it is on, but the model carries
whatever the author set — an unset field is omitted so a round-trip never fabricates one. An
empty object emits neither element.

```ts
interface PageSetup {
    fitToPage?: boolean;
    fitToWidth?: number;
    fitToHeight?: number;
    scale?: number;
    orientation?: 'portrait' | 'landscape';
    pageOrder?: 'downThenOver' | 'overThenDown';
    paperSize?: number;
    printerSettings?: Uint8Array;
}
```

---

### `PrintOptions`

<sub>interface</sub>

Print-toggle flags from the `<printOptions>` element. Each maps to a boolean OOXML attribute that
defaults false — except `gridLinesSet`, which defaults true and gates whether `gridLines` is
honoured. The model stores only what the source or caller set, so an unset flag is omitted and a
round-trip never fabricates one; an empty object emits no element at all.

```ts
interface PrintOptions {
    horizontalCentered?: boolean;
    verticalCentered?: boolean;
    headings?: boolean;
    gridLines?: boolean;
    gridLinesSet?: boolean;
}
```
