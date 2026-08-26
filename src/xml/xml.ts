// Minimal XML serialisation helpers for the write path.
//
// Writing OOXML needs only correct escaping and well-formed structure; parsing (the
// reader's concern) is a separate, later decision, so no XML library is on the write
// path. Escaping is the one hard, security-relevant requirement, since an unescaped `<`,
// `&`, or `"` produces a malformed package a consumer rejects, so it lives here,
// audited once, rather than sprinkled through the part emitters.
//
// A second class of character is harder than `&`: the ones XML 1.0 has no syntax for at
// all. `&#1;` is not an escape for U+0001, it is another way of spelling the same illegal
// document. So there are two behaviours here, and which one applies is decided by the
// destination rather than by the character.
//
// **Prose the user typed** is escaped `_xHHHH_`. That is SpreadsheetML's own convention, and
// where it applies Excel both writes it and decodes it on the way back in. Which elements are
// in that group is a measured fact, not a schema one, and each was settled over COM against
// Excel Desktop: a `<t>` inline string, the cached `<v>` of a `t="str"` formula cell, and a
// threaded comment's `<text>`, which all read back as the decoded character and are written
// back out re-escaped by Excel itself. So `textElement`, the `t="str"` result and the
// threaded-comment body get `escapeSpreadsheetText`.
//
// **Everything else is refused.** The schema is no help in drawing that line: `ST_Xstring`
// types a sheet name and a print header just as it types `<t>`, so following the type alone
// would have us escape a sheet name too, and a workbook whose tab reads `Sheet_x0001_A` is
// not a faithful rendering of the name the author asked for, it is a different name. Where
// there is no faithful representation the honest answer is `AuthoringError`, which is the
// stance `numberText` below already takes on a non-finite number for the same reason. That
// covers every structural string: sheet and defined names, formulas, table column names,
// document properties, relationship targets, and a print header's text.
//
// The two comment systems land on the same side of that line by two different routes, which
// is why the test is *does Excel decode it here* rather than *does this read as prose*. A
// legacy note's body is a `<t>` in a `CT_Rst`, the very type the convention is defined on. A
// threaded comment's `<text>` is a different element in the 2018 extension namespace with no
// documented escape at all; it took a measurement to put it here, and it decodes with the
// same closed grammar and the same single pass. A print header stays refused: nobody has
// measured a decode there, and its own `&`-prefixed formatting codes are the only in-band
// syntax it has.
//
// Refusing is a new throw on a path that used to "succeed" by producing a file Excel reports
// as damaged, so the failure moved earlier and got louder, which is the whole trade.

import {AuthoringError} from '../errors.ts';

const TEXT_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
};

const ATTR_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
  '\n': '&#10;',
  '\r': '&#13;',
  '\t': '&#9;',
};

/**
 * The characters an XML 1.0 document cannot carry, whatever escape you reach for.
 *
 * Three classes: the C0 controls outside the tab/LF/CR the `Char` production allows, the two
 * noncharacters at the top of the BMP, and unpaired surrogates. The last are not an XML
 * problem but a UTF-8 one: the encoder substitutes U+FFFD for a lone surrogate, so the
 * package validates and the value is quietly gone, which is the same loss by a different
 * route. U+007F and the C1 controls are deliberately absent: XML 1.1 forbids them, OOXML is
 * 1.0.
 *
 * The `u` flag is load-bearing. It makes the pattern match code points, so an astral
 * character is one unit that no surrogate range can match, and `[\u{D800}-\u{DFFF}]` means
 * exactly "a surrogate that is not part of a pair" with no lookaround.
 */
// oxlint-disable-next-line eslint/no-control-regex -- naming the control characters is the point: this pattern exists to find them
const UNREPRESENTABLE = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;
const UNREPRESENTABLE_GLOBAL = new RegExp(UNREPRESENTABLE.source, 'gu');

/** `U+0001`-style spelling of a code point, for an escape body or an error message. */
function codePointHex(codePoint: number): string {
  return codePoint.toString(16).toUpperCase().padStart(4, '0');
}

/**
 * Refuse a string that XML cannot carry, naming the character and where it is so the author
 * can find it in a value they never inspected. These arrive from a database column or a CSV
 * field, not from a literal in the calling code.
 *
 * Exported for the one escape that cannot be {@link escapeAttr}: a number format code escapes
 * everything but the apostrophe, and needs this guard just the same.
 *
 * @throws {AuthoringError} naming the code point and its offset.
 */
export function assertRepresentable(value: string): void {
  const found = UNREPRESENTABLE.exec(value);
  if (found === null) return;
  throw new AuthoringError(
    `cannot write U+${codePointHex(value.codePointAt(found.index) as number)} at offset ${found.index}: ` +
      'XML 1.0 has no representation for it, and the _xHHHH_ escape that would carry it is a ' +
      'convention of cell values only',
  );
}

/** Escape a string for use as XML element text. */
export function escapeText(value: string): string {
  assertRepresentable(value);
  return value.replace(/[&<>]/g, (ch) => TEXT_ESCAPES[ch] as string);
}

/** Escape a string for use inside a double-quoted XML attribute value. */
export function escapeAttr(value: string): string {
  assertRepresentable(value);
  return value.replace(/[&<>"'\n\r\t]/g, (ch) => ATTR_ESCAPES[ch] as string);
}

/**
 * Escape a cell value: the `_xHHHH_` convention first, then XML's own escapes.
 *
 * The order of the two replacements is load-bearing. Once `_xHHHH_` decodes to something, a
 * value that legitimately reads `_x0041_` would come back as `A`, so the underscore is
 * escaped first, and only where it begins a sequence that would otherwise decode, leaving
 * every other underscore alone. Doing that after the character escape would re-escape the
 * `_x005F_` it had just introduced.
 *
 * `escapeText` runs last and sees no unrepresentable character left, so its guard is a
 * standing proof that the escape was complete rather than a second check of the same thing.
 *
 * Reversed by `decodeSpreadsheetText` in `./xml-read.ts`, whose single left-to-right pass is what
 * makes the `_x005F_` step above reversible. Change either and read its comment first.
 */
export function escapeSpreadsheetText(value: string): string {
  return escapeText(
    value
      .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
      .replace(UNREPRESENTABLE_GLOBAL, (ch) => `_x${codePointHex(ch.codePointAt(0) as number)}_`),
  );
}

/**
 * Whether an element's text must be wrapped with `xml:space="preserve"` to survive a
 * round-trip. Leading/trailing whitespace is otherwise collapsed by consumers, so a
 * string cell value that begins or ends with a space needs the marker.
 */
function needsSpacePreserve(value: string): boolean {
  return value.length > 0 && (value !== value.trim() || /[\n\r\t]/.test(value));
}

/**
 * A `<t>` text element carrying an escaped string, marked `xml:space="preserve"` when its
 * whitespace would otherwise be collapsed. Shared by every string-bearing element (a plain
 * inline string cell, a rich-text run, a pooled string, a note's body) so all decode
 * identically on the way back.
 */
export function textElement(value: string): string {
  const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : '';
  return `<t${space}>${escapeSpreadsheetText(value)}</t>`;
}

/**
 * Render a formula operand for serialisation: a number becomes its literal, a string is stripped of
 * the single optional leading '=' an author may write (OOXML stores the expression without it, e.g.
 * `=A1>0` on disk is `A1>0`). The result is unescaped: the caller escapes it for its target,
 * whether that is element text or an attribute value.
 */
export function stripFormulaEquals(value: string | number): string {
  if (typeof value === 'number') return String(value);
  return value.startsWith('=') ? value.slice(1) : value;
}

/**
 * A boolean attribute rendered with a leading space (` name="1"` / ` name="0"`), or '' when the value
 * is undefined. OOXML booleans serialise as 1/0; emitting the explicit `="0"` lets a writer force a
 * flag off against a consumer's default, while an unset (undefined) flag stays out of the element
 * entirely: the two-state-plus-absent contract every flag writer here shares.
 */
export function boolAttr(name: string, value: boolean | undefined): string {
  return value === undefined ? '' : ` ${name}="${value ? 1 : 0}"`;
}

/**
 * A numeric attribute rendered with a leading space (` name="42"`), or '' when the value is undefined
 * so that a count an author never set stays out of the element rather than fabricating a default.
 */
export function attr(name: string, value: number | undefined): string {
  return value === undefined ? '' : ` ${name}="${value}"`;
}

/**
 * A finite number serialises as its shortest round-trippable decimal; a non-finite one
 * has no OOXML numeric representation, so the writer refuses it rather than emit `NaN`.
 */
export function numberText(value: number): string {
  if (!Number.isFinite(value)) {
    throw new AuthoringError(
      `cannot write a non-finite number (${value}): it has no OOXML representation`,
    );
  }
  return String(value);
}

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
