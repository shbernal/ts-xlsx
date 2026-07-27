// The workbook theme's colour scheme — the twelve colours a `theme="n"` reference resolves against.
//
// The theme part itself is carried opaquely by the model (see `Workbook.restoreThemePart`); this
// module reads just the `<a:clrScheme>` out of it, because that is the only piece a colour reference
// needs. Everything else in a theme (the format scheme's gradients, line and effect styles) is
// nobody's business here.

/**
 * The twelve colour-scheme slots **in the order a `theme="n"` attribute indexes them**.
 *
 * This order is not the order the slots appear in the theme part. ISO/IEC 29500 §20.1.6.2 documents
 * the `<a:clrScheme>` child sequence as `dk1, lt1, dk2, lt2, accent1…6, hlink, folHlink`, and that is
 * how the XML is written — but SpreadsheetML's `theme="n"` does **not** index that sequence. Excel
 * swaps each dark/light pair: index 0 is `lt1`, 1 is `dk1`, 2 is `lt2`, 3 is `dk2`.
 *
 * Verified against Excel Desktop rather than inferred, because the two orders differ only in the
 * first four entries and reading either one into the other silently inverts text against background
 * — see `docs/knowledge/specs/theme-color-index-order.md` and the recorded observation in
 * `test/corpus/fixtures/excel-oracle/theme-color-index-order.json`. The stylesheet's own default font
 * is the everyday witness: it carries `<color theme="1"/>` and renders black, which is `dk1`.
 */
export const THEME_COLOR_SLOTS = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/** One slot of a theme's colour scheme. */
export type ThemeColorSlot = (typeof THEME_COLOR_SLOTS)[number];

/**
 * A theme's colour scheme: each slot's colour as a 6-hex `RRGGBB` string. Partial because a foreign
 * theme is free to omit a slot (or express one in a colour model this reader does not decode), and an
 * absent slot is honestly absent rather than silently substituted.
 */
export type ThemeColorScheme = Readonly<Partial<Record<ThemeColorSlot, string>>>;

/** The Office default colour scheme, matching the theme part the writer emits for a workbook with none. */
export const DEFAULT_THEME_COLOR_SCHEME: ThemeColorScheme = {
  lt1: 'FFFFFF',
  dk1: '000000',
  lt2: 'E7E6E6',
  dk2: '44546A',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

// One `<a:slot>` of a `<a:clrScheme>` and the colour element inside it. Two colour models appear in
// practice: `<a:srgbClr val="RRGGBB"/>` states the colour directly, while `<a:sysClr val="windowText"
// lastClr="000000"/>` defers to an operating-system colour and records what it last resolved to.
// `dk1`/`lt1` are almost always the sysClr form, so a reader that only understands srgbClr resolves
// nothing for the two most-referenced slots in any workbook.
const SCHEME_SLOT = new RegExp(
  '<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)\\b[^>]*>' + '\\s*<a:(srgbClr|sysClr)\\b([^>]*)>',
  'g',
);

/**
 * Extract the colour scheme from a theme part. Returns only the slots the part actually declares in a
 * colour model this reader understands; an unrecognised one is dropped rather than guessed at, so a
 * caller can tell "the theme says nothing here" from "the theme says black".
 *
 * Reads the `<a:clrScheme>` block alone. A theme carries a font scheme and a format scheme too, but
 * neither participates in resolving a colour, and scanning the whole part would let a `<a:srgbClr>`
 * buried in a gradient stop masquerade as a scheme slot.
 */
export function parseThemeColorScheme(themeXml: string): ThemeColorScheme {
  const block = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml);
  if (block === null) return {};
  const scheme: {-readonly [K in ThemeColorSlot]?: string} = {};
  for (const match of (block[1] ?? '').matchAll(SCHEME_SLOT)) {
    const slot = match[1] as ThemeColorSlot;
    const attrs = match[3] ?? '';
    // A sysClr's `val` is a system-colour name ("windowText"), not a colour — its `lastClr` is the
    // concrete value the authoring application last resolved that name to, and is the only thing here
    // a consumer without the same OS theme can use.
    const source = match[2] === 'sysClr' ? /\blastClr="([^"]*)"/ : /\bval="([^"]*)"/;
    const value = source.exec(attrs)?.[1];
    if (value !== undefined && /^[0-9a-fA-F]{6}$/.test(value)) scheme[slot] = value;
  }
  return scheme;
}
