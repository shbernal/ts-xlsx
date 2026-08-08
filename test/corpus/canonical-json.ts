// Comparing two objects that mean the same thing but were built by different parsers.
//
// A corpus case very often asserts that two paths produce the *same* structure — a workbook read from
// XML against the same workbook read from the binary form, a fixture before a round-trip against the
// fixture after. `JSON.stringify` cannot express that, because key order is an artefact of which
// parser filled the object and in what order it happened to assign: two structurally identical
// results serialise to two different strings. Sorting keys at every depth first removes that artefact
// and nothing else.
//
// It was written three separate times inside the adapter — twice as `stableSort` and once as
// `canonical`, each `any` in and `any` out — which is how it ended up with three slightly different
// spellings of the same recursion and no single place to state what it guarantees.

/**
 * `value` with every object's keys sorted, recursively. Arrays keep their order; primitives, `null`
 * and anything else pass through untouched.
 *
 * Intended to be handed straight to `JSON.stringify`, so that two structurally equal values compare
 * equal as strings regardless of the order their keys were assigned in. It is **not** a
 * general-purpose deep clone: a `Date`, a `Map`, or a class instance is returned as-is and will
 * serialise however `JSON.stringify` chooses to serialise it.
 */
export function canonicalJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((element) => canonicalJson(element));
  if (value === null || typeof value !== 'object') return value;
  // `typeof value === 'object'` narrows to `object`, which has no index signature — the cast states
  // the one thing the narrowing cannot: that string-keyed access is what `Object.keys` just promised.
  const record = value as Record<string, unknown>;
  const sorted: Record<string, unknown> = {};
  for (const key of Object.keys(record).sort()) sorted[key] = canonicalJson(record[key]);
  return sorted;
}
