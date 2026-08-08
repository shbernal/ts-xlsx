// The adapter's remaining type debt, named so it can be counted.
//
// `Untyped` is `any`. It exists because one name used to mean two opposite things: `CorpusApi` was
// both "the adapter surface a case calls" and "a value somewhere in the adapter we never got around
// to typing", and the second meaning is what kept the first one from ever becoming real. Splitting
// them lets the surface be typed today and leaves the debt visible: `grep -c Untyped` is the balance,
// and every occurrence is a place where a wrong shape reaches runtime instead of the compiler.
//
// Prefer a real type. Reach for this only when the value's shape is genuinely owned by a case (a
// declarative spec, an options bag a case invented) rather than by `src`, and even then prefer
// writing the shape down. It is not a license, it is an IOU.
//
// It lives in its own module with no imports so the typed surface (`CorpusApi`, derived from the
// adapter) and this can never tangle into an import cycle.

/** A value the adapter has not typed yet. See the module comment — this is debt, not a design. */
// biome-ignore lint/suspicious/noExplicitAny: that is the entire point of this alias.
export type Untyped = any;
