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
