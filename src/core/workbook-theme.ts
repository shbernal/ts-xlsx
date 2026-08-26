// The workbook's theme, lifted off `Workbook` into the slice it actually is.
//
// `theme.ts` holds the theme *model*: the slot order, the scheme parsers, `applyThemeOverrides`.
// What lives here is the per-workbook state and caching wrapped around that model: the preserved
// part, the decoded-scheme cache, and the overrides a caller authored. Folding this into `theme.ts`
// would mix a pure model module with mutable per-instance cache, which is how `theme.ts` grows the
// same legibility problem `Workbook` had.
//
// The boundary that matters, and the reason this slice is not as clean as the VBA one: colour
// resolution needs the workbook's custom `<indexedColors>` palette as well as the theme scheme, and
// that palette is *also* the writer's source for the `<indexedColors>` element and is populated by
// the reader. So it stays on `Workbook` and reaches here as a narrow accessor. Moving it would drag
// the styles-table state along and turn a theme overlay into a colour-and-styles overlay, which is
// not a slice.
//
// The doc comments for the public surface stay on `Workbook`'s accessors, which is what the API
// reference is generated from and what a consumer reads.

import {resolveColor} from './color-resolution.ts';
import type {Color} from './style.ts';
import {
  applyThemeOverrides,
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  DEFAULT_THEME_XML,
  parseThemeColorScheme,
  parseThemeFontScheme,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from './theme.ts';
import type {PreservedTheme} from './workbook.ts';

/**
 * The theme slice of a workbook: the preserved part, the colour scheme and typefaces every
 * `theme="n"` reference resolves against, and the overrides {@link Workbook.setTheme} authors over
 * them.
 */
export class WorkbookTheme {
  // The workbook's custom indexed palette, read on demand rather than held: it belongs to the styles
  // state, not to the theme, and only colour resolution needs it.
  readonly #indexedPalette: () => readonly string[];

  // The theme part read from a file, kept verbatim with the closure of parts it reaches. The writer
  // emits its own default theme for a workbook that has none, so without this a branded theme would be
  // overwritten by that default and every `theme="n"` colour in the file would silently re-render.
  // Undefined for a workbook authored from scratch, or read from a package declaring no theme.
  #part: PreservedTheme | undefined;

  // The theme's colour scheme, decoded from the preserved part (and merged with any authored
  // overrides) on first use. Cached because resolving a colour is a per-cell operation and the part is
  // otherwise held as bytes; invalidated whenever the theme is replaced or authored. The two events
  // are `restorePart` and `author`, and there is nowhere else that can stale it.
  #colors: ThemeColorScheme | undefined;

  // Colour slots and typefaces the caller authored, merged over whatever the workbook already had.
  #authored: {colors: {-readonly [K in ThemeColorSlot]?: string}; fonts: ThemeFontScheme} = {
    colors: {},
    fonts: {},
  };

  constructor(indexedPalette: () => readonly string[]) {
    this.#indexedPalette = indexedPalette;
  }

  get part(): PreservedTheme | undefined {
    return this.#part;
  }

  // The reader's channel: a restored part invalidates the decoded scheme.
  restorePart(theme: PreservedTheme | undefined): void {
    this.#part = theme;
    this.#colors = undefined;
  }

  author(overrides: ThemeOverrides): void {
    // Validated eagerly, by running the generation the writer will later run: a colour rejected at
    // write time would surface far from the call that supplied it.
    applyThemeOverrides(this.#baseXml(), overrides);
    Object.assign(this.#authored.colors, overrides.colors ?? {});
    this.#authored.fonts = {...this.#authored.fonts, ...overrides.fonts};
    this.#colors = undefined;
  }

  get colors(): ThemeColorScheme {
    if (this.#colors === undefined) {
      const xml = this.#xml();
      // A theme that declares no scheme (or none this reader decodes) falls back to the Office
      // default rather than resolving nothing: the file still renders against *some* scheme, and the
      // default is the one the writer would have shipped.
      const parsed = xml === undefined ? {} : parseThemeColorScheme(xml);
      const base = Object.keys(parsed).length === 0 ? DEFAULT_THEME_COLOR_SCHEME : parsed;
      this.#colors = {...base, ...this.#authored.colors};
    }
    return this.#colors;
  }

  get fonts(): ThemeFontScheme {
    const xml = this.#xml();
    const parsed = xml === undefined ? {} : parseThemeFontScheme(xml);
    const base = Object.keys(parsed).length === 0 ? DEFAULT_THEME_FONTS : parsed;
    return {...base, ...this.#authored.fonts};
  }

  // The typefaces a caller authored, unmerged. `Workbook.defaultFont` needs to know whether the body
  // face was named outright or merely inherited, which the resolved {@link fonts} cannot say.
  get authoredFonts(): ThemeFontScheme {
    return this.#authored.fonts;
  }

  authoredXml(): string | undefined {
    const {colors, fonts} = this.#authored;
    if (Object.keys(colors).length === 0 && Object.keys(fonts).length === 0) return undefined;
    return applyThemeOverrides(this.#baseXml(), {colors, fonts});
  }

  resolveColor(color: Color): string | undefined {
    return resolveColor(color, {theme: this.colors, indexed: this.#indexedPalette()});
  }

  // The part authored overrides are applied on top of: the preserved source theme, else the default
  // one the writer would otherwise have emitted.
  #baseXml(): string {
    return this.#xml() ?? DEFAULT_THEME_XML;
  }

  // The preserved theme part's text, decoded from the entry part of its closure.
  #xml(): string | undefined {
    const theme = this.#part;
    if (theme === undefined) return undefined;
    const entry = theme.parts.find((part) => part.path === theme.entryPath);
    return entry === undefined ? undefined : new TextDecoder().decode(entry.bytes);
  }
}
