# Indexed-palette colors must resolve to a concrete color, not surface a raw index

Cluster: styles

## Scenario

A caller inspects a cell's fill (or font/border) color and finds, instead of a usable `argb` value, a
bare integer `indexed` reference — e.g. a `bgColor` carrying `indexed="64"` with no RGB. They have no
way to turn that number into an actual color. This happens because OOXML permits several color
encodings — direct ARGB, theme-color-plus-tint, and a legacy 64-entry **indexed palette** carried
over from the old BIFF/`.xls` format — and the library hands back the raw indexed reference without
resolving it. Files produced by older tools, or by converters that emit the legacy palette, are full
of these.

> Spec note, not a corpus case: the underlying report is a support question with no bug or fixture,
> but it names a real gap in the color model. The durable value is the resolution contract; a concrete
> assertion becomes possible once the model resolves indexed colors (a fixture using an `indexed`
> fill, asserted to expose a concrete color), building on the existing indexed-palette round-trip case.

## Desired behavior

- **Every color the public model exposes is resolvable to a concrete ARGB** without the caller knowing
  the encoding. An `indexed` reference resolves through the palette (the standard 64-entry default, or
  a workbook-supplied `<indexedColors>` override) to an ARGB value the caller can read and render.
- **The encoding is not lost.** Resolving to ARGB for consumption must not erase the fact that the
  source was indexed (or themed) — a color carries both its resolved ARGB and its original reference,
  so a round-trip can re-emit the same encoding rather than rewriting every color as literal ARGB and
  bloating the styles table / changing the file's semantics.
- **The default palette is built in.** The legacy 64-entry palette (including the special
  `indexed="64"`/`65` "automatic" foreground/background system colors) is known to the library, so a
  file that references it without shipping a custom `<indexedColors>` block still resolves correctly.
- **A workbook-level custom indexed palette is honored.** When the styles part declares its own
  `<indexedColors>`, those entries override the defaults for that workbook on both read and write.

## Open questions

- The public shape of a resolved color: always expose `argb` plus an optional `{indexed}`/`{theme,tint}`
  origin, or a discriminated union the caller narrows? The types are the docs — this must be precise.
- Whether writing a color the caller supplied as ARGB should ever be down-converted to an indexed entry
  (probably never — indexed is a read-tolerance and round-trip-fidelity concern, not an authoring one).
- How this composes with theme-color resolution (same "resolve but remember the origin" principle) and
  the `indexed="64"` automatic/system-color special cases.

Related: `custom-indexed-color-palette-roundtrip`, `theme-color-font-backed-by-theme-part`,
`public-type-surface-matches-runtime`.
