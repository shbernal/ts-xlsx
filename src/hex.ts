// A number as fixed-width uppercase hexadecimal.
//
// One line of code, and it had six copies -- an error message naming a code point, a byte dump, a
// colour channel, an image digest -- which had already diverged on both of the two axes it has: the
// pad width, and whether the digits are uppercased. Divergence on a formatter is invisible until the
// day two of its outputs are compared, and one of these is a *key*.
//
// Its own module rather than a corner of `bytes.ts`, because `errors.ts` is the one consumer that
// must stay cheap: `/errors` is the entry a service imports to classify a failure without paying for
// a parser, and pulling in the byte primitives for one line would have doubled it. This imports
// nothing, so every layer can reach it.

/**
 * `value` in uppercase hexadecimal, zero-padded to at least `digits`.
 *
 * Uppercase because five of the six sites this replaced were, because OOXML writes `ARGB` values that
 * way, and because a hex digit read by a human beside a decimal number is easier to tell apart when
 * it is not lowercase. A value wider than `digits` is not truncated: padding is a minimum.
 */
export function hex(value: number, digits: number): string {
  return value.toString(16).toUpperCase().padStart(digits, '0');
}
