# Hyperlink cells get the built-in link appearance by default

Cluster: styles

## Scenario

A user writes cells whose value is a hyperlink (text + target URL, optionally a tooltip) — by
assigning a hyperlink-shaped value, adding rows containing hyperlink values, or via the
streaming writer. In native spreadsheet applications a link cell renders in the standard blue
underlined "Hyperlink" look and switches to the "Followed Hyperlink" (purple) look after being
clicked. The library emits the correct link relationship and clickable target, but the cell
carries no link styling, so the text is ordinary black un-underlined and never changes color
when visited. Users manually apply an underline + blue font to every link cell, which
approximates the look but loses the visited-state color change.

## Desired behavior

When the library writes a cell that carries a hyperlink, the cell should by default receive the
standard built-in **Hyperlink** appearance — underline plus the **theme's hyperlink color
(theme color index 10)** — so it renders like a native link and gains the followed/visited
state. A hardcoded blue font is insufficient: it does not reproduce the visited-color behavior,
and it does not follow the theme when the theme changes.

## Open questions

- **Default vs opt-in.** A default that mimics native applications is friendlier but must not
  silently clobber a font the user already set. Reasonable rule: apply the default hyperlink
  font only when the cell has no explicit font override; an explicit font wins.
- **Named cell style vs raw font.** Native applications use a named `Hyperlink` /
  `Followed Hyperlink` cellStyle in `styles.xml`; that is what unlocks the automatic
  visited-state color. A raw font override only fakes the unvisited look. Investigate whether
  emitting the named cell style is required for the visited transition, or whether a
  theme-color-10 underline font alone suffices.
- **Streaming parity.** The same defaulting must apply on the streaming write path, not only the
  in-memory path.
- **Entry-point convergence.** Hyperlinks set via `addRow`/`addRows` rich-value shapes and via
  direct cell assignment should all converge on the same styling.

## Prior art / workarounds observed

- Manual per-cell font `{ underline: true, color: { argb: 'FF0000FF' } }` — approximates the
  unvisited look, no visited state.
- Confirmed better: `{ ...cell.font, underline: true, color: { theme: 10 } }` — uses the theme
  hyperlink color and yields the working visited state.

## Assertable behavior (for a future corpus case, once the design is settled)

A written cell containing a hyperlink, with no explicit font, serializes with an underlined font
referencing theme color index 10 (or the emitted named Hyperlink cell style); a cell with an
explicit font retains that font.
