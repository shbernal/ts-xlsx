// Minimal XML serialisation helpers for the write path.
//
// Writing OOXML needs only correct escaping and well-formed structure; parsing (the
// reader's concern) is a separate, later decision, so no XML library is on the write
// path. Escaping is the one hard, security-relevant requirement — an unescaped `<`,
// `&`, or `"` produces a malformed package a consumer rejects — so it lives here,
// audited once, rather than sprinkled through the part emitters.

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

/** Escape a string for use as XML element text. */
export function escapeText(value: string): string {
  return value.replace(/[&<>]/g, (ch) => TEXT_ESCAPES[ch] as string);
}

/** Escape a string for use inside a double-quoted XML attribute value. */
export function escapeAttr(value: string): string {
  return value.replace(/[&<>"'\n\r\t]/g, (ch) => ATTR_ESCAPES[ch] as string);
}

/**
 * Whether an element's text must be wrapped with `xml:space="preserve"` to survive a
 * round-trip. Leading/trailing whitespace is otherwise collapsed by consumers, so a
 * string cell value that begins or ends with a space needs the marker.
 */
export function needsSpacePreserve(value: string): boolean {
  return value.length > 0 && (value !== value.trim() || /[\n\r\t]/.test(value));
}

/**
 * A `<t>` text element carrying an escaped string, marked `xml:space="preserve"` when its
 * whitespace would otherwise be collapsed. Shared by every string-bearing element — a plain
 * inline string cell, a rich-text run — so all decode identically on the way back.
 */
export function textElement(value: string): string {
  const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : '';
  return `<t${space}>${escapeText(value)}</t>`;
}

/**
 * Render a formula operand for serialisation: a number becomes its literal, a string is stripped of
 * the single optional leading '=' an author may write (OOXML stores the expression without it, e.g.
 * `=A1>0` on disk is `A1>0`). The result is unescaped — the caller escapes it for its target,
 * whether that is element text or an attribute value.
 */
export function stripFormulaEquals(value: string | number): string {
  if (typeof value === 'number') return String(value);
  return value.startsWith('=') ? value.slice(1) : value;
}

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
