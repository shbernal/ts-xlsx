// Which code points an XML 1.0 document may carry, stated once, for both directions of the wire.
//
// The rule belongs to neither half of `src/xml/`. The writer needs it to refuse a value it cannot
// serialise; the reader needs it to refuse a value it would otherwise decode into the model and hand
// straight to that refusal, several layers and one taxonomy boundary later. Putting it in `xml.ts`
// would make the reader import the serialisation vocabulary and `AuthoringError` with it, which is
// the exact coupling `xml-scan.ts` was split out to avoid; putting it in `xml-scan.ts` would make the
// writer import the scanner. So it lives below both and imports nothing.
//
// One declaration, two shapes. The regex is the statement of the rule and the predicate is derived
// from it, because the alternative -- a range list written out twice -- is a rule that can be half
// changed.

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
export const XML_UNREPRESENTABLE =
  // oxlint-disable-next-line eslint/no-control-regex -- naming the control characters is the point: this pattern exists to find them
  /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;

/** The same rule with the `g` flag, for the escapes that rewrite every occurrence. */
export const XML_UNREPRESENTABLE_GLOBAL = new RegExp(XML_UNREPRESENTABLE.source, 'gu');

/**
 * Whether a code point may appear in an XML 1.0 document at all.
 *
 * The reader's question, and the reason this module exists. A numeric character reference names a
 * code point directly, so a file is free to name one the format has no representation for; decoding
 * it anyway puts a value in the model that the writer is *guaranteed* to refuse, turning a hostile
 * file into an `AuthoringError` blaming the caller on the next save.
 *
 * A code point outside Unicode entirely is not this function's business -- the caller has already
 * bounded the number before it can name a character -- so an out-of-range argument is refused here
 * as well rather than throwing out of `String.fromCodePoint`.
 */
export function isRepresentableCodePoint(codePoint: number): boolean {
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) return false;
  return !XML_UNREPRESENTABLE.test(String.fromCodePoint(codePoint));
}

/**
 * Drop every code point XML 1.0 cannot carry.
 *
 * For text the scanner read out of a document, where there is no verbatim form to fall back to: a
 * *reference* to an unrepresentable character can be left as the `&#1;` the file wrote, but a raw
 * one has no spelling of its own to keep. Such a character makes the document ill-formed by the
 * `Char` production, so it was never legally in the file; admitting it would put a value in the
 * model that the writer must then refuse, which reports a corrupt input as the caller's mistake.
 */
export function stripUnrepresentable(text: string): string {
  if (!XML_UNREPRESENTABLE.test(text)) return text;
  return text.replace(XML_UNREPRESENTABLE_GLOBAL, '');
}
