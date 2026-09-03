// Narrowing guards for the closed enumerations OOXML spells as attribute tokens.
//
// A reader meets these tokens as bare strings out of a foreign file, so each union the model
// declares needs a predicate saying whether a given string is one of its members. Written by hand,
// that predicate pairs a lookup table with a `value is Union` return type derived from nothing: a
// guard reading the neighbouring table compiles, passes its tests, and narrows to the wrong union.
// Deriving the guard from the table removes that hazard by construction: the table *is* the type.

/**
 * Build a narrowing predicate for a closed token union.
 *
 * `members` is keyed by the union, so the compiler refuses a foreign member and an omitted one
 * alike. A bare list of string literals is checked one way only: it cannot name a token the union
 * does not, but it can quietly omit one the union does, and a guard narrower than its union
 * silently drops a token real files legitimately carry.
 *
 * Deriving the union from the list instead (`(typeof TOKENS)[number]`, as `table-style.ts` does)
 * would make divergence impossible rather than merely detected, and is the better shape for a type
 * that is not public. It is not available to a published union: the API reference renders a
 * declaration by slicing its own source text, so a derived alias reaches the docs as that
 * expression instead of as its members.
 */
export function tokenSet<T extends string>(
  members: Record<T, true>,
): (value: string) => value is T {
  return tokenSetOf(Object.keys(members) as T[]);
}

/**
 * {@link tokenSet} for a union whose members are also needed as an ordered *list*, because a binary
 * serialisation indexes one.
 *
 * The record form's whole guarantee is that a member cannot be omitted, and a list gives that up: it
 * can name only members of the union, but it can name fewer. So a caller here owes an explicit proof
 * of set equality in both directions, which is what `FILL_PATTERNS_IN_SCHEMA_ORDER` carries beside
 * itself. That is a worse guarantee than the record's, taken deliberately: the alternative is the
 * arrangement it replaced, where the union, the guard's table and the binary codec's index list were
 * three hand-maintained copies of one enumeration and only two of them were checked against each
 * other.
 */
export function tokenSetOf<T extends string>(members: readonly T[]): (value: string) => value is T {
  const set = new Set<string>(members);
  return (value): value is T => set.has(value);
}
