# Xlsx Theme Xml

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DEFAULT_THEME_XML`

<sub>const</sub>

The theme part a workbook with no theme of its own ships: the standard Office theme.

A spreadsheet must carry one even when nobody configured it: the stylesheet's own default font
references `theme="1"`, which a consumer can only resolve against this part, so the two travel
together. It is also the base `applyThemeOverrides` authors on top of when a workbook was
built from scratch rather than read from a file.

```ts
const DEFAULT_THEME_XML: string
```

---

### `parseThemeColorScheme`

<sub>function</sub>

Extract the colour scheme from a theme part. Returns only the slots the part actually declares in a
colour model this reader understands; an unrecognised one is dropped rather than guessed at, so a
caller can tell "the theme says nothing here" from "the theme says black".

Reads the `<clrScheme>` block alone. A theme carries a font scheme and a format scheme too, but
neither participates in resolving a colour, and scanning the whole part would let an `<srgbClr>`
buried in a gradient stop masquerade as a scheme slot.

```ts
function parseThemeColorScheme(themeXml: string): ThemeColorScheme;
```

---

### `parseThemeFontScheme`

<sub>function</sub>

Extract the major/minor latin typefaces from a theme part's `<fontScheme>`.

```ts
function parseThemeFontScheme(themeXml: string): ThemeFontScheme;
```
