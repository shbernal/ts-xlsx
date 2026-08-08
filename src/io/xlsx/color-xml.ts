// The `<color>` element, both directions.
//
// OOXML spells a colour the same way wherever one appears — `<color>`, `<fgColor>`, `<bgColor>`,
// `<tabColor>`, a differential format's colour — as some combination of `rgb`, `theme`, `tint` and
// `indexed` attributes. Writing that tuple and reading it back are one concern with two directions,
// and they have to agree: `parseColor` must accept exactly what `colorAttrs` emits, or a round-trip
// loses a colour.
//
// They lived in `styles.ts`, the write-side style *table*, which made the style-table reader and the
// worksheet reader import from the writer to decode a colour — the only place in this codec where the
// read pipeline reached into the write pipeline. Nothing about decoding `<color>` belongs to the
// interning table; it just happened to be where the first caller was.

import type {Color} from '../../core/style.ts';
import {AuthoringError} from '../../errors.ts';

// OOXML wants a bare 8-hex ARGB (alpha + RGB). This single choke point — through which every
// fill/font/border/tab colour flows — accepts two developer conveniences and rejects the rest loudly,
// because a malformed rgb value does not error in Excel: it silently renders as flat black.
//   - A leading '#' is a CSS habit and is stripped ('#FFBFBFBF' → 'FFBFBFBF').
//   - A 6-hex RGB is promoted to ARGB with a fully-opaque alpha ('00FF00' → 'FF00FF00'), the common
//     case of a colour written without its alpha channel.
// Anything not then exactly 8 hex digits is a programming error at the API surface, so it throws with
// the offending value rather than writing corrupt XML. Casing is preserved so foreign files round-trip.
function normalizeArgb(argb: string): string {
  const hex = argb.startsWith('#') ? argb.slice(1) : argb;
  const rgb = hex.length === 6 ? `FF${hex}` : hex;
  if (!/^[0-9a-fA-F]{8}$/.test(rgb)) {
    throw new AuthoringError(
      `Invalid ARGB colour ${JSON.stringify(argb)}: expected 6 or 8 hexadecimal digits`,
    );
  }
  return rgb;
}

/** Serialise a {@link Color} as the attribute list a `<color>`-shaped element carries. */
export function colorAttrs(color: Color): string {
  const parts: string[] = [];
  if (color.argb !== undefined) parts.push(`rgb="${normalizeArgb(color.argb)}"`);
  if (color.theme !== undefined) parts.push(`theme="${color.theme}"`);
  if (color.tint !== undefined) parts.push(`tint="${color.tint}"`);
  if (color.indexed !== undefined) parts.push(`indexed="${color.indexed}"`);
  return parts.join(' ');
}

// The read counterpart of colorAttrs: decode a `<color>`/`<fgColor>`/… element's attributes.
// theme/indexed must be integers and tint a finite number; a malformed foreign attribute is dropped
// rather than propagated as NaN, so a downstream colorAttrs never emits `theme="NaN"`.
export function parseColor(attrs: {readonly [k: string]: string}): Color {
  const color: {argb?: string; theme?: number; tint?: number; indexed?: number} = {};
  if (attrs.rgb !== undefined) color.argb = attrs.rgb;
  if (attrs.theme !== undefined) {
    const theme = Number(attrs.theme);
    if (Number.isInteger(theme)) color.theme = theme;
  }
  if (attrs.tint !== undefined) {
    const tint = Number(attrs.tint);
    if (Number.isFinite(tint)) color.tint = tint;
  }
  if (attrs.indexed !== undefined) {
    const indexed = Number(attrs.indexed);
    if (Number.isInteger(indexed)) color.indexed = indexed;
  }
  return color;
}
