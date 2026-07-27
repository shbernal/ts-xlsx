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

/**
 * The two typefaces a theme nominates: the `major` face headings use and the `minor` face body text
 * uses. A cell's font reaches them by `scheme="major"`/`scheme="minor"` instead of naming a typeface,
 * so changing these restyles every such cell at once.
 */
export interface ThemeFontScheme {
  readonly major?: string | undefined;
  readonly minor?: string | undefined;
}

/** The Office default typefaces, matching the theme part the writer emits for a workbook with none. */
export const DEFAULT_THEME_FONTS: ThemeFontScheme = {major: 'Calibri Light', minor: 'Calibri'};

/** Extract the major/minor latin typefaces from a theme part's `<a:fontScheme>`. */
export function parseThemeFontScheme(themeXml: string): ThemeFontScheme {
  const block = /<a:fontScheme\b[^>]*>([\s\S]*?)<\/a:fontScheme>/.exec(themeXml);
  if (block === null) return {};
  const face = (which: 'majorFont' | 'minorFont'): string | undefined => {
    const font = new RegExp(`<a:${which}\\b[^>]*>([\\s\\S]*?)</a:${which}>`).exec(block[1] ?? '');
    return /<a:latin\b[^>]*\btypeface="([^"]*)"/.exec(font?.[1] ?? '')?.[1];
  };
  const scheme: {major?: string; minor?: string} = {};
  const major = face('majorFont');
  const minor = face('minorFont');
  if (major !== undefined) scheme.major = major;
  if (minor !== undefined) scheme.minor = minor;
  return scheme;
}

/** What a caller can author on a workbook's theme: any subset of the colour slots and typefaces. */
export interface ThemeOverrides {
  readonly colors?: Readonly<Partial<Record<ThemeColorSlot, string>>> | undefined;
  readonly fonts?: ThemeFontScheme | undefined;
}

// The `<a:clrScheme>` child order — dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink. Not the order
// `theme="n"` indexes (see THEME_COLOR_SLOTS); this is the sequence CT_ColorScheme requires the
// elements to be written in, and writing them in index order would be schema-invalid.
const SCHEME_ELEMENT_ORDER: readonly ThemeColorSlot[] = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/**
 * Apply authored colour/font overrides to a theme part, returning the new part text.
 *
 * Surgical by design: the base part rides through untouched except for the `<a:clrScheme>` and
 * `<a:fontScheme>` blocks, and within those, only what the caller actually named. The format scheme —
 * the gradients, line styles and effect styles that make a theme look like a theme — is left exactly
 * as it was, because nobody hand-authors `fillStyleLst` gradient stops from a spreadsheet API and
 * regenerating it would replace a designer's work with the Office default.
 *
 * A slot the caller did not override keeps its **verbatim source element**, not a re-serialisation of
 * its value. That matters for `dk1`/`lt1`, which Excel writes as `<a:sysClr val="windowText"
 * lastClr="000000"/>`: rewriting those as `<a:srgbClr>` would pin them to one machine's resolved
 * window colours and break dark-mode following.
 */
export function applyThemeOverrides(baseXml: string, overrides: ThemeOverrides): string {
  let xml = baseXml;
  const colors = overrides.colors ?? {};
  if (Object.keys(colors).length > 0) {
    const sourceElements = parseThemeColorElements(baseXml);
    const body = SCHEME_ELEMENT_ORDER.map((slot) => {
      const authored = colors[slot];
      const inner =
        authored !== undefined
          ? `<a:srgbClr val="${normalizeSchemeValue(authored)}"/>`
          : (sourceElements[slot] ??
            `<a:srgbClr val="${DEFAULT_THEME_COLOR_SCHEME[slot] as string}"/>`);
      return `<a:${slot}>${inner}</a:${slot}>`;
    }).join('');
    xml = replaceBlockBody(xml, 'clrScheme', body);
  }
  const {major, minor} = overrides.fonts ?? {};
  if (major !== undefined) xml = replaceLatinTypeface(xml, 'majorFont', major);
  if (minor !== undefined) xml = replaceLatinTypeface(xml, 'minorFont', minor);
  return xml;
}

/**
 * Each colour slot's verbatim inner element from a theme part — `<a:srgbClr val="…"/>` or
 * `<a:sysClr val="…" lastClr="…"/>`. The value-level counterpart is {@link parseThemeColorScheme};
 * this keeps the *encoding* so an untouched slot can be re-emitted exactly as the source wrote it.
 */
export function parseThemeColorElements(
  themeXml: string,
): Readonly<Partial<Record<ThemeColorSlot, string>>> {
  const block = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml);
  if (block === null) return {};
  const elements: {-readonly [K in ThemeColorSlot]?: string} = {};
  const pattern =
    /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>([\s\S]*?)<\/a:\1>|<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)\/>/g;
  for (const match of (block[1] ?? '').matchAll(pattern)) {
    const slot = (match[1] ?? match[3]) as ThemeColorSlot;
    const inner = match[2];
    if (inner !== undefined && inner !== '') elements[slot] = inner;
  }
  return elements;
}

// A theme colour is a bare 6-hex RGB — DrawingML has no alpha channel on `<a:srgbClr val>`. The two
// conveniences the rest of the library accepts (a leading '#', an 8-hex ARGB) are accepted and
// reduced here; anything else is a caller's bug and is refused rather than written as corrupt XML,
// which Excel does not report — it renders the slot as flat black.
function normalizeSchemeValue(value: string): string {
  const hex = value.startsWith('#') ? value.slice(1) : value;
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9a-fA-F]{6}$/.test(rgb)) {
    throw new Error(
      `Invalid theme colour ${JSON.stringify(value)}: expected 6 hexadecimal digits (RRGGBB)`,
    );
  }
  return rgb.toUpperCase();
}

// Replace the body of `<a:name>…</a:name>`, keeping the element's own attributes (the scheme's
// display name). A base with no such block is left alone: this authors an existing theme, and a theme
// that declares no colour scheme at all is not one an override can repair.
function replaceBlockBody(xml: string, name: string, body: string): string {
  const pattern = new RegExp(`(<a:${name}\\b[^>]*>)[\\s\\S]*?(</a:${name}>)`);
  return xml.replace(pattern, (_all, open: string, close: string) => `${open}${body}${close}`);
}

// Swap just the `<a:latin typeface="…"/>` inside one of the two font slots, leaving its `panose` and
// the east-asian/complex-script faces beside it as they were.
function replaceLatinTypeface(
  xml: string,
  which: 'majorFont' | 'minorFont',
  typeface: string,
): string {
  const pattern = new RegExp(`(<a:${which}\\b[^>]*>[\\s\\S]*?<a:latin\\b)[^>]*(/>)`);
  return xml.replace(
    pattern,
    (_all, open: string, close: string) => `${open} typeface="${escapeXmlAttr(typeface)}"${close}`,
  );
}

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * The theme part a workbook with no theme of its own ships — the standard Office theme.
 *
 * A spreadsheet must carry one even when nobody configured it: the stylesheet's own default font
 * references `theme="1"`, which a consumer can only resolve against this part, so the two travel
 * together. It is also the base {@link applyThemeOverrides} authors on top of when a workbook was
 * built from scratch rather than read from a file.
 */
export const DEFAULT_THEME_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">' +
  '<a:themeElements>' +
  '<a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme>' +
  '<a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
  '</a:fontScheme>' +
  '<a:fmtScheme name="Office">' +
  '<a:fillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
  '</a:fillStyleLst>' +
  '<a:lnStyleLst>' +
  '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
  '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
  '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
  '</a:lnStyleLst>' +
  '<a:effectStyleLst>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
  '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
  '</a:bgFillStyleLst>' +
  '</a:fmtScheme>' +
  '</a:themeElements>' +
  '<a:objectDefaults/>' +
  '<a:extraClrSchemeLst/>' +
  '</a:theme>';
