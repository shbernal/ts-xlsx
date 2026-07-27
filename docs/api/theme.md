# Theme

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DEFAULT_THEME_COLOR_SCHEME`

<sub>const</sub>

The Office default colour scheme, matching the theme part the writer emits for a workbook with none.

```ts
const DEFAULT_THEME_COLOR_SCHEME: Readonly<Partial<Record<"lt1" | "dk1" | "lt2" | "dk2" | "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink", string>>>
```

---

### `parseThemeColorScheme`

<sub>function</sub>

Extract the colour scheme from a theme part. Returns only the slots the part actually declares in a
colour model this reader understands; an unrecognised one is dropped rather than guessed at, so a
caller can tell "the theme says nothing here" from "the theme says black".

Reads the `<a:clrScheme>` block alone. A theme carries a font scheme and a format scheme too, but
neither participates in resolving a colour, and scanning the whole part would let a `<a:srgbClr>`
buried in a gradient stop masquerade as a scheme slot.

```ts
function parseThemeColorScheme(themeXml: string): ThemeColorScheme;
```

---

### `THEME_COLOR_SLOTS`

<sub>const</sub>

The twelve colour-scheme slots **in the order a `theme="n"` attribute indexes them**.

This order is not the order the slots appear in the theme part. ISO/IEC 29500 §20.1.6.2 documents
the `<a:clrScheme>` child sequence as `dk1, lt1, dk2, lt2, accent1…6, hlink, folHlink`, and that is
how the XML is written — but SpreadsheetML's `theme="n"` does **not** index that sequence. Excel
swaps each dark/light pair: index 0 is `lt1`, 1 is `dk1`, 2 is `lt2`, 3 is `dk2`.

Verified against Excel Desktop rather than inferred, because the two orders differ only in the
first four entries and reading either one into the other silently inverts text against background
— see `docs/knowledge/specs/theme-color-index-order.md` and the recorded observation in
`test/corpus/fixtures/excel-oracle/theme-color-index-order.json`. The stylesheet's own default font
is the everyday witness: it carries `<color theme="1"/>` and renders black, which is `dk1`.

```ts
const THEME_COLOR_SLOTS: readonly ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
```

---

### `ThemeColorScheme`

<sub>type</sub>

A theme's colour scheme: each slot's colour as a 6-hex `RRGGBB` string. Partial because a foreign
theme is free to omit a slot (or express one in a colour model this reader does not decode), and an
absent slot is honestly absent rather than silently substituted.

```ts
type ThemeColorScheme = Readonly<Partial<Record<ThemeColorSlot, string>>>;
```

---

### `ThemeColorSlot`

<sub>type</sub>

One slot of a theme's colour scheme.

```ts
type ThemeColorSlot = (typeof THEME_COLOR_SLOTS)[number];
```
