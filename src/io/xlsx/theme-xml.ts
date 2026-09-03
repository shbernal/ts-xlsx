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
import {parseXml} from '../../xml/xml-read.ts';
import {localName, type XmlAttributes} from '../../xml/xml-scan.ts';
import {escapeAttr} from '../../xml/xml.ts';

/**
 * Everything the three theme readers want, from one scan.
 *
 * On the SAX reader rather than on container-scanning regular expressions. `xml-read.ts` names that
 * anti-pattern and says why it was removed: "a regular expression parsing XML, over untrusted input,
 * in the same directory as the reader written specifically to avoid that (ADR-0004)". This file was
 * the caller that was missed, and it had the failure mode to go with it: every pattern hardcoded the
 * `a:` prefix, so a theme in the *default* DrawingML namespace, or under any other prefix, read as no
 * theme at all. The workbook then fell back to the library's own theme and every theme-indexed colour
 * in the file resolved to the wrong RGB, with no error anywhere. An XML comment between a slot and its
 * colour child broke the same patterns, which required whitespace and nothing else between them.
 *
 * `elements` keeps each slot's colour child as source text so an override can re-emit an untouched
 * slot exactly as it arrived, rebuilt from the element's own qualified name and attributes rather
 * than sliced out of the part: the child is always a single empty element, so the two are the same
 * bytes, and rebuilding is what keeps the prefix the *source* used instead of the one we assume.
 */
interface ThemeScheme {
  readonly colors: ThemeColorScheme;
  readonly elements: Readonly<Partial<Record<ThemeColorSlot, string>>>;
  readonly fonts: ThemeFontScheme;
  /** The prefix the part binds DrawingML to, `''` when it is the default namespace. */
  readonly prefix: string;
}

// The colour models a scheme slot states its value with. `sysClr` defers to an operating-system colour
// and records what it last resolved to, and `dk1`/`lt1` are almost always that form, so a reader that
// understands only `srgbClr` resolves nothing for the two most-referenced slots in any workbook.
const COLOR_ELEMENTS = new Set(['srgbClr', 'sysClr']);

function readThemeScheme(themeXml: string): ThemeScheme {
  const colors: {-readonly [K in ThemeColorSlot]?: string} = {};
  const elements: {-readonly [K in ThemeColorSlot]?: string} = {};
  const fonts: {major?: string; minor?: string} = {};
  let prefix = '';

  // Which container the scan is inside. A theme carries a format scheme too, and a `<srgbClr>` buried
  // in one of its gradient stops must not be read as a scheme slot, which is why matching is scoped to
  // a container rather than done over the whole part.
  let container: 'clrScheme' | 'fontScheme' | undefined;
  let slot: ThemeColorSlot | undefined;
  let fontSlot: 'major' | 'minor' | undefined;

  parseXml(themeXml, {
    onOpen(name, attrs, _selfClosing, scope) {
      const local = localName(name);
      if (local === 'theme') {
        const colon = name.indexOf(':');
        prefix = colon === -1 ? '' : name.slice(0, colon + 1);
        return;
      }
      if (local === 'clrScheme' || local === 'fontScheme') {
        container = local;
        return;
      }
      if (container === 'clrScheme') {
        if (isThemeColorSlot(local)) {
          slot = local;
        } else if (slot !== undefined && COLOR_ELEMENTS.has(local)) {
          // A sysClr's `val` is a system-colour name ("windowText"), not a colour. Its `lastClr` is
          // the concrete value the authoring application last resolved that name to, and is the only
          // thing here a consumer without the same OS theme can use.
          const value = local === 'sysClr' ? attrs.lastClr : attrs.val;
          if (value !== undefined && /^[0-9a-fA-F]{6}$/.test(value)) colors[slot] = value;
          elements[slot] = emptyElement(name, attrs);
          slot = undefined;
        }
        return;
      }
      if (container === 'fontScheme') {
        if (local === 'majorFont') fontSlot = 'major';
        else if (local === 'minorFont') fontSlot = 'minor';
        else if (local === 'latin' && fontSlot !== undefined && attrs.typeface !== undefined) {
          // The attribute arrives already entity-decoded from the scanner, which is what the regex
          // reader had to do by hand and what made authoring "A&B" read back as "A&amp;B" before it.
          fonts[fontSlot] ??= attrs.typeface;
        }
      }
      void scope;
    },
    onClose(name) {
      const local = localName(name);
      if (local === 'clrScheme' || local === 'fontScheme') {
        container = undefined;
        slot = undefined;
        fontSlot = undefined;
      } else if (local === 'majorFont' || local === 'minorFont') {
        fontSlot = undefined;
      } else if (slot !== undefined && local === slot) {
        slot = undefined;
      }
    },
  });

  const scheme: {major?: string; minor?: string} = {};
  if (fonts.major !== undefined) scheme.major = fonts.major;
  if (fonts.minor !== undefined) scheme.minor = fonts.minor;
  return {colors, elements, fonts: scheme, prefix};
}

// Re-render an empty element from its parsed name and attributes. Attribute order is the scanner's,
// which is the source's, so for the single-element colour children this captures it is the source text.
function emptyElement(name: string, attrs: XmlAttributes): string {
  let out = `<${name}`;
  for (const key in attrs) out += ` ${key}="${escapeAttr(attrs[key] ?? '')}"`;
  return `${out}/>`;
}

/**
 * Extract the colour scheme from a theme part. Returns only the slots the part actually declares in a
 * colour model this reader understands; an unrecognised one is dropped rather than guessed at, so a
 * caller can tell "the theme says nothing here" from "the theme says black".
 *
 * Reads the `<clrScheme>` block alone. A theme carries a font scheme and a format scheme too, but
 * neither participates in resolving a colour, and scanning the whole part would let an `<srgbClr>`
 * buried in a gradient stop masquerade as a scheme slot.
 */
export function parseThemeColorScheme(themeXml: string): ThemeColorScheme {
  return readThemeScheme(themeXml).colors;
}

/** Extract the major/minor latin typefaces from a theme part's `<fontScheme>`. */
export function parseThemeFontScheme(themeXml: string): ThemeFontScheme {
  return readThemeScheme(themeXml).fonts;
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
  // The prefix the base part itself binds DrawingML to, so what is written back matches the part
  // being edited rather than the one this library happens to ship. Under any prefix but `a`, every
  // replacement below used to match nothing and the overrides were silently discarded.
  const {elements: sourceElements, prefix} = readThemeScheme(baseXml);
  const colors = overrides.colors ?? {};
  if (Object.keys(colors).length > 0) {
    const body = SCHEME_ELEMENT_ORDER.map((slot) => {
      const authored = colors[slot];
      const inner =
        authored !== undefined
          ? `<${prefix}srgbClr val="${normalizeThemeColor(authored)}"/>`
          : (sourceElements[slot] ??
            `<${prefix}srgbClr val="${DEFAULT_THEME_COLOR_SCHEME[slot]}"/>`);
      return `<${prefix}${slot}>${inner}</${prefix}${slot}>`;
    }).join('');
    xml = replaceBlockBody(xml, prefix, 'clrScheme', body);
  }
  const {major, minor} = overrides.fonts ?? {};
  if (major !== undefined) xml = replaceLatinTypeface(xml, prefix, 'majorFont', major);
  if (minor !== undefined) xml = replaceLatinTypeface(xml, prefix, 'minorFont', minor);
  return xml;
}

// Replace the body of `<clrScheme>…</clrScheme>`, keeping the element's own attributes (the scheme's
// display name). A base with no such block is left alone: this authors an existing theme, and a theme
// that declares no colour scheme at all is not one an override can repair.
function replaceBlockBody(xml: string, prefix: string, name: string, body: string): string {
  const pattern = new RegExp(`(<${prefix}${name}\\b[^>]*>)[\\s\\S]*?(</${prefix}${name}>)`);
  return xml.replace(pattern, (_all, open: string, close: string) => `${open}${body}${close}`);
}

// Swap just the `<latin typeface="…"/>` inside one of the two font slots, leaving its `panose` and
// the east-asian/complex-script faces beside it as they were.
function replaceLatinTypeface(
  xml: string,
  prefix: string,
  which: 'majorFont' | 'minorFont',
  typeface: string,
): string {
  const pattern = new RegExp(`(<${prefix}${which}\\b[^>]*>[\\s\\S]*?<${prefix}latin\\b)[^>]*(/>)`);
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
