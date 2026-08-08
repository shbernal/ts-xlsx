// The adapter's remaining type debt, named so it can be counted.
//
// `Untyped` is `any`. It exists because one name used to mean two opposite things: `CorpusApi` was
// both "the adapter surface a case calls" and "a value somewhere in the adapter we never got around
// to typing", and the second meaning is what kept the first one from ever becoming real. Splitting
// them lets the surface be typed today and leaves the debt visible — every occurrence is a place
// where a wrong shape reaches runtime instead of the compiler. The balance is
//
//   grep -ro Untyped test/corpus --include=*.ts | wc -l
//
// with `-o` because `grep -c` counts *lines*, and several of these sit two to a line.
//
// What is left is deliberate, and knowing which is which saves the next reader a search. A **case
// spec** — the declarative `{sheets: [{cells: [...]}]}` a case hands `buildFrom`, and the options bags
// built around it — is genuinely owned by the cases, and writing its type down is a real piece of
// work rather than an annotation. A **report accumulator** (`Record<string, Untyped>`) is the other
// half: widening it to `unknown` does not type anything, it just moves the debt into the 254 cases
// that read the report.
//
// Everything else has been paid off. In particular, a value whose type `src` already publishes is
// *not* on that list: `runtime.ts` exports `WorkbookInstance`, `WorksheetInstance` and `CellInstance`
// for exactly that, and reaching for `Untyped` instead switches off the checking the corpus is for.
//
// Prefer a real type. It is not a license, it is an IOU.
//
// It lives in its own module with no imports so the typed surface (`CorpusApi`, derived from the
// adapter) and this can never tangle into an import cycle.

/** A value the adapter has not typed yet. See the module comment — this is debt, not a design. */
// biome-ignore lint/suspicious/noExplicitAny: that is the entire point of this alias.
export type Untyped = any;
