// The workbook theme, as the model knows it: the twelve colour slots a `theme="n"` reference
// resolves against, the two typefaces a `scheme="major"`/`scheme="minor"` font resolves to, and what
// a caller may author over either.
//
// Values only. How a theme is *written down* -- the `<a:theme>` part, the readers that pull these
// schemes out of one, and the surgery that applies an override onto one -- is a serialisation, and
// lives with the codec that carries it (`io/xlsx/theme-xml.ts`). The part itself rides through the
// model opaquely (see `Workbook.restoreThemePart`); nothing here parses it.

/**
 * The twelve colour-scheme slots **in the order a `theme="n"` attribute indexes them**.
 *
 * This order is not the order the slots appear in the theme part. ISO/IEC 29500 §20.1.6.2 documents
 * the `<a:clrScheme>` child sequence as `dk1, lt1, dk2, lt2, accent1…6, hlink, folHlink`, and that is
 * how the XML is written. But SpreadsheetML's `theme="n"` does **not** index that sequence: Excel
 * swaps each dark/light pair: index 0 is `lt1`, 1 is `dk1`, 2 is `lt2`, 3 is `dk2`.
 *
 * Verified against Excel Desktop rather than inferred, because the two orders differ only in the
 * first four entries and reading either one into the other silently inverts text against
 * background. See `docs/knowledge/specs/theme-color-index-order.md` and the recorded observation in
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

const THEME_COLOR_SLOT_NAMES: ReadonlySet<string> = new Set<string>(THEME_COLOR_SLOTS);

/** Narrow a raw `<a:clrScheme>` child name to a known {@link ThemeColorSlot}. */
export function isThemeColorSlot(value: string): value is ThemeColorSlot {
  return THEME_COLOR_SLOT_NAMES.has(value);
}

/**
 * A theme's colour scheme: each slot's colour as a 6-hex `RRGGBB` string. Partial because a foreign
 * theme is free to omit a slot (or express one in a colour model this reader does not decode), and an
 * absent slot is honestly absent rather than silently substituted.
 */
export type ThemeColorScheme = Readonly<Partial<Record<ThemeColorSlot, string>>>;

/** The Office default colour scheme, matching the theme part the writer emits for a workbook with none. */
// Typed as a complete record rather than a `ThemeColorScheme`: this scheme declares every slot, and
// saying so is what lets a consumer index it without an assertion. A slot added to the union without
// a colour here is a compile error, which is the point.
export const DEFAULT_THEME_COLOR_SCHEME: Readonly<Record<ThemeColorSlot, string>> = {
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

/**
 * The two typefaces a theme nominates: the `major` face headings use and the `minor` face body text
 * uses. A cell's font reaches them by `scheme="major"`/`scheme="minor"` instead of naming a typeface,
 * so changing these restyles every such cell at once.
 */
export interface ThemeFontScheme {
  readonly major?: string | undefined;
  readonly minor?: string | undefined;
}

/**
 * The body typeface a workbook falls back to when neither its theme nor its styles part names one:
 * the face the default theme nominates, and so the face every `scheme="minor"` font resolves to.
 * Named rather than inlined because it is also the last resort of the default-font chain
 * ({@link Workbook.defaultFont}), and the two must not drift.
 */
export const OFFICE_BODY_FACE = 'Calibri';

/** The Office default typefaces, matching the theme part the writer emits for a workbook with none. */
export const DEFAULT_THEME_FONTS: ThemeFontScheme = {
  major: 'Calibri Light',
  minor: OFFICE_BODY_FACE,
};

/** What a caller can author on a workbook's theme: any subset of the colour slots and typefaces. */
export interface ThemeOverrides {
  readonly colors?: Readonly<Partial<Record<ThemeColorSlot, string>>> | undefined;
  readonly fonts?: ThemeFontScheme | undefined;
}

/** Reduce an authored theme colour to the bare 6-hex RGB DrawingML wants.
 *
 * A theme colour is a bare 6-hex RGB: DrawingML has no alpha channel on `<a:srgbClr val>`. The two
 * conveniences the rest of the library accepts (a leading '#', an 8-hex ARGB) are accepted and
 * reduced here; anything else is a caller's bug and is refused rather than written as corrupt XML,
 * which Excel does not report; it renders the slot as flat black.
 *
 * @throws {SyntaxError} if the value is not a recognisable RGB or ARGB hex string. Native rather
 * than the library's own `AuthoringError`: a string that does not parse is what `SyntaxError` is
 * for, and it is the same kind of failure a malformed comment GUID raises.
 */
export function normalizeThemeColor(value: string): string {
  const hex = value.startsWith('#') ? value.slice(1) : value;
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9a-fA-F]{6}$/.test(rgb)) {
    throw new SyntaxError(
      `Invalid theme colour ${JSON.stringify(value)}: expected 6 hexadecimal digits (RRGGBB)`,
    );
  }
  return rgb.toUpperCase();
}
