// The `<font>` wire form, both directions.
//
// A stylesheet's `<font>` and a rich-text run's `<rPr>` are the same element in two contexts: they
// share every child, and differ only in the face, which CT_Font names `<name>` and CT_RPrElt names
// `<rFont>`. Both halves already handled both spellings, so this was one wire form already; what it
// lacked was a module.
//
// It had lived inside the 744-line style-table reader, so the shared-strings path and the inline-string
// path each dragged that whole reader in for a font. `color-xml.ts` is the precedent this follows:
// small, both directions, imported by both sides, and neither side importing the other.

import {
  type Font,
  isFontScheme,
  isFontVerticalAlignment,
  isNamedUnderlineStyle,
  type UnderlineStyle,
} from '../../core/style.ts';
import {numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {boolPresent, type XmlAttributes} from '../../xml/xml-scan.ts';
import {checkedToken, escapeAttr, numberText} from '../../xml/xml.ts';
import {colorAttrs, parseColor} from './color-xml.ts';

// A mutable font accumulator while a <font> element's children stream in; frozen into a
// Font on close.
export type FontDraft = {-readonly [K in keyof Font]?: Font[K]};

// A <font> child element sets one facet on the draft. Boolean flags honour their `val`: a
// bare tag or val="1"/"true" is on, val="0"/"false" is off (an explicit-false flag is not
// truthy merely because the tag is present). An unrecognised child is ignored.
export function applyFontChild(draft: FontDraft, local: string, attrs: XmlAttributes): void {
  switch (local) {
    case 'b':
      draft.bold = boolPresent(attrs.val);
      break;
    case 'i':
      draft.italic = boolPresent(attrs.val);
      break;
    case 'strike':
      draft.strike = boolPresent(attrs.val);
      break;
    case 'outline':
      draft.outline = boolPresent(attrs.val);
      break;
    case 'u':
      // A bare <u/> is a single underline; a named style (single/double/…) carries through; but
      // val="none" is the explicit ABSENCE of an underline, so it must read back falsy, not the
      // truthy string "none" that a consumer's `if (font.underline)` would mistake for underlined. An
      // unrecognised token keeps the "is underlined" fact but drops the unknown style (a plain true).
      draft.underline =
        attrs.val === undefined
          ? true
          : attrs.val === 'none'
            ? false
            : isNamedUnderlineStyle(attrs.val)
              ? attrs.val
              : true;
      break;
    case 'vertAlign':
      if (attrs.val !== undefined && isFontVerticalAlignment(attrs.val))
        draft.vertAlign = attrs.val;
      break;
    case 'sz': {
      const size = numFinite(attrs.val);
      if (size !== undefined) draft.size = size;
      break;
    }
    case 'color':
      draft.color = parseColor(attrs);
      break;
    // `<name>` in a styles `<font>`, `<rFont>` in a rich-text run's `<rPr>`: the same font face.
    case 'name':
    case 'rFont':
      if (attrs.val !== undefined) draft.name = attrs.val;
      break;
    case 'family': {
      const family = numInteger(attrs.val);
      if (family !== undefined) draft.family = family;
      break;
    }
    case 'charset': {
      const charset = numInteger(attrs.val);
      if (charset !== undefined) draft.charset = charset;
      break;
    }
    case 'scheme':
      if (attrs.val !== undefined && isFontScheme(attrs.val)) draft.scheme = attrs.val;
      break;
    default:
      break;
  }
}

// Serialise the facets a font overrides, in ECMA-376 child order. A boolean flag is emitted only
// when true (its absence is the default false); an empty result means the font differs from the
// default in nothing and needs no entry at all. The face element differs by context: a styles
// `<font>` names it `<name>` (CT_Font) and a rich-text run's `<rPr>` names it `<rFont>` (CT_RPrElt).
// Otherwise the two share every child, so `nameTag` selects the face element and the rest is common.
export function fontXml(font: Font, nameTag: 'name' | 'rFont' = 'name'): string {
  const parts: string[] = [];
  if (font.bold) parts.push('<b/>');
  if (font.italic) parts.push('<i/>');
  if (font.strike) parts.push('<strike/>');
  if (font.outline) parts.push('<outline/>');
  const underline = underlineXml(font.underline);
  if (underline !== '') parts.push(underline);
  if (font.vertAlign !== undefined) {
    parts.push(
      `<vertAlign val="${checkedToken(font.vertAlign, isFontVerticalAlignment, 'font vertical alignment')}"/>`,
    );
  }
  if (font.size !== undefined) parts.push(`<sz val="${numberText(font.size)}"/>`);
  if (font.color !== undefined) parts.push(`<color ${colorAttrs(font.color)}/>`);
  if (font.name !== undefined) parts.push(`<${nameTag} val="${escapeAttr(font.name)}"/>`);
  if (font.family !== undefined) parts.push(`<family val="${numberText(font.family)}"/>`);
  if (font.charset !== undefined) parts.push(`<charset val="${numberText(font.charset)}"/>`);
  if (font.scheme !== undefined && font.scheme !== 'none')
    parts.push(`<scheme val="${checkedToken(font.scheme, isFontScheme, 'font scheme')}"/>`);
  return parts.join('');
}

// `<u/>` is single underline (the same as an explicit "single"); the named variants carry a
// val; false and "none" are the default no-underline and emit nothing.
function underlineXml(underline: UnderlineStyle | undefined): string {
  if (underline === undefined || underline === false || underline === 'none') return '';
  if (underline === true || underline === 'single') return '<u/>';
  return `<u val="${checkedToken(underline, isNamedUnderlineStyle, 'underline style')}"/>`;
}
