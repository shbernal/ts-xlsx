# Color Resolution

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `applyTint`

<sub>function</sub>

Apply a `tint` to a concrete ARGB colour: convert to HLS, shift the luminance, convert back.

The shift is ECMA-376's (§18.3.1.15): `-1` darkens to black, `+1` lightens to white, `0` is a
no-op. The spec states the formula but not the rounding either side of the RGB↔HLS conversion, so
an implementation cannot be checked against the prose alone. This one lands within 2/255 per
channel of what Excel Desktop itself renders, measured over three hues × twelve tints and recorded
in `test/corpus/fixtures/excel-oracle/theme-color-tint-luminance.json`. Closing that last gap would
mean reproducing one Excel build's fixed-point HLS rounding, and the difference is not visible.

```ts
function applyTint(argb: string, tint: number): string;
```

---

### `ColorResolutionContext`

<sub>interface</sub>

What a colour reference is resolved against: the workbook's theme scheme and indexed palette.

```ts
interface ColorResolutionContext {
  /** The theme's colour scheme; the Office default when the workbook carries no theme of its own. */
  readonly theme?: ThemeColorScheme | undefined;
  /**
   * The workbook's custom indexed palette, by index, each entry an ARGB string. Empty or absent means
   * the workbook rides {@link DEFAULT_INDEXED_COLORS}. A custom palette replaces the built-in one
   * wholesale — that is what `<indexedColors>` means — so a short custom palette leaves the indices
   * past its end unresolved rather than falling through to the built-in entry.
   */
  readonly indexed?: readonly string[] | undefined;
}
```

---

### `DEFAULT_INDEXED_COLORS`

<sub>const</sub>

The built-in indexed colour palette (ECMA-376 §18.8.27), by index. Entries 0–7 duplicate 8–15 —
redundancy the spec preserves for backwards compatibility with the legacy formats this palette came
from — and the table is only 64 long: indices 64 and 65 are the *system* foreground and background,
which name whatever the operating system's window colours are and therefore have no fixed value at
all (see [`SYSTEM_INDEXED_COLORS`](./color-resolution.md#systemindexedcolors)).

The spec writes each entry with a leading `00`. That byte is not an alpha channel — a palette of
fully transparent colours would be absurd — it is an artefact of the 32-bit colour records these
values were lifted from, which is why [`resolveColor`](./color-resolution.md#resolvecolor) returns them fully opaque.

```ts
const DEFAULT_INDEXED_COLORS: readonly string[]
```

---

### `resolveColor`

<sub>function</sub>

Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it cannot be
resolved — an `auto` colour, a system indexed colour, a theme slot the workbook's scheme does not
declare, or an index past the end of a custom palette.

Precedence follows what the encodings mean: an explicit `argb` is already concrete and wins; then
`theme`, then `indexed`. A `tint` applies to whatever the base resolved to.

Alpha: a resolved `theme`/`indexed` colour comes back fully opaque, because neither the theme scheme
nor the palette carries a meaningful alpha (see [`DEFAULT_INDEXED_COLORS`](./color-resolution.md#defaultindexedcolors)). An explicit `argb`
keeps the alpha the file stated.

```ts
function resolveColor(
  color: Color,
  context: ColorResolutionContext = {},
): string | undefined;
```

---

### `SYSTEM_INDEXED_COLORS`

<sub>const</sub>

The two indices that are not colours: 64 is the system foreground and 65 the system background.
They resolve to whatever the viewing system's window colours are, so this library reports them
unresolved rather than inventing black and white — a caller that wants to paint them must decide
for itself what "automatic" means in its context. `indexed="64"` in particular is extremely common:
it is the placeholder every solid fill Excel writes carries as its background colour.

```ts
const SYSTEM_INDEXED_COLORS: ReadonlySet<number>
```
