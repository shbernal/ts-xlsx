// The theme part, both directions.
//
// `<a:theme>` is DrawingML, not SpreadsheetML, but a workbook cannot do without one: the stylesheet's
// own default font references `theme="1"`, so a consumer can only resolve it against this part. Here
// live the part the writer ships when a workbook has none, the two readers that pull a colour scheme
// and a font scheme back out of one, and the surgery that applies authored overrides onto an existing
// part without disturbing anything the caller did not name.
//
// It is here and not in `core/` because it *is* a serialisation, which is the one thing the model
// layer is defined as not having (`scripts/check-layering.ts`). While it lived there it could not
// import `src/xml/`, and paid for that with a hand-rolled attribute escape weaker than the real one.
// `color-xml.ts` is the precedent for a read-and-write-in-one-file feature module in this directory.
// Nothing but the xlsx codec reads a theme today; the day a second one does, this moves to a shared
// home and not before.

import {
  DEFAULT_THEME_COLOR_SCHEME,
  isThemeColorSlot,
  normalizeThemeColor,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from '../../core/theme.ts';
import {decodeEntities} from '../../xml/xml-scan.ts';
import {escapeAttr} from '../../xml/xml.ts';

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
    const slot = match[1];
    if (slot === undefined || !isThemeColorSlot(slot)) continue;
    const attrs = match[3] ?? '';
    // A sysClr's `val` is a system-colour name ("windowText"), not a colour. Its `lastClr` is the
    // concrete value the authoring application last resolved that name to, and is the only thing here
    // a consumer without the same OS theme can use.
    const source = match[2] === 'sysClr' ? /\blastClr="([^"]*)"/ : /\bval="([^"]*)"/;
    const value = source.exec(attrs)?.[1];
    if (value !== undefined && /^[0-9a-fA-F]{6}$/.test(value)) scheme[slot] = value;
  }
  return scheme;
}

/** Extract the major/minor latin typefaces from a theme part's `<a:fontScheme>`. */
export function parseThemeFontScheme(themeXml: string): ThemeFontScheme {
  const block = /<a:fontScheme\b[^>]*>([\s\S]*?)<\/a:fontScheme>/.exec(themeXml);
  if (block === null) return {};
  const face = (which: 'majorFont' | 'minorFont'): string | undefined => {
    const font = new RegExp(`<a:${which}\\b[^>]*>([\\s\\S]*?)</a:${which}>`).exec(block[1] ?? '');
    const typeface = /<a:latin\b[^>]*\btypeface="([^"]*)"/.exec(font?.[1] ?? '')?.[1];
    // Decoded, unlike the colour scheme's slots above: those are six hex digits and can carry no
    // entity, but a typeface is free text and the writer's `escapeAttr` turns an `&` in one into
    // `&amp;`. Handing the raw attribute back is what made authoring "A&B" read back as "A&amp;B".
    return typeface === undefined ? undefined : decodeEntities(typeface);
  };
  const scheme: {major?: string; minor?: string} = {};
  const major = face('majorFont');
  const minor = face('minorFont');
  if (major !== undefined) scheme.major = major;
  if (minor !== undefined) scheme.minor = minor;
  return scheme;
}

// The `<a:clrScheme>` child order: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink. Not the order
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
 * `<a:fontScheme>` blocks, and within those, only what the caller actually named. The format scheme,
 * the gradients, line styles and effect styles that make a theme look like a theme, is left exactly
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
          ? `<a:srgbClr val="${normalizeThemeColor(authored)}"/>`
          : (sourceElements[slot] ?? `<a:srgbClr val="${DEFAULT_THEME_COLOR_SCHEME[slot]}"/>`);
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
 * Each colour slot's verbatim inner element from a theme part: `<a:srgbClr val="…"/>` or
 * `<a:sysClr val="…" lastClr="…"/>`. The value-level counterpart is {@link parseThemeColorScheme};
 * this keeps the *encoding* so an untouched slot can be re-emitted exactly as the source wrote it.
 */
function parseThemeColorElements(
  themeXml: string,
): Readonly<Partial<Record<ThemeColorSlot, string>>> {
  const block = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml);
  if (block === null) return {};
  const elements: {-readonly [K in ThemeColorSlot]?: string} = {};
  const pattern =
    /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>([\s\S]*?)<\/a:\1>|<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)\/>/g;
  for (const match of (block[1] ?? '').matchAll(pattern)) {
    const slot = match[1] ?? match[3];
    if (slot === undefined || !isThemeColorSlot(slot)) continue;
    const inner = match[2];
    if (inner !== undefined && inner !== '') elements[slot] = inner;
  }
  return elements;
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
    (_all, open: string, close: string) => `${open} typeface="${escapeAttr(typeface)}"${close}`,
  );
}

/**
 * The theme part a workbook with no theme of its own ships: the standard Office theme.
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
