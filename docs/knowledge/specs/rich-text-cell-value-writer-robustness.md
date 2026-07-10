# The writer must not crash on a malformed rich-text-shaped cell value

Cluster: styles

## Scenario

A user builds a workbook programmatically and assigns a cell a value that is, or resembles, a
rich-text value (an object with a `richText` runs array). On save, the writer walks every cell and
reads the runs array off each value. When a cell's value is a rich-text-shaped object whose runs
array is missing or undefined — an incomplete object built up in application code, or a partially
constructed value — a legacy serializer dereferenced the runs array unguarded and threw a deep-stack
`TypeError: Cannot read properties of undefined (reading 'richText')`, with no indication of which
cell caused it.

> Spec note, not a corpus case: probing the current writer shows a `{richText: undefined}` value does
> NOT crash it, so the exact malformed shape that triggers the throw is version/shape-specific and
> not reliably reproducible. The durable requirement — the writer must never crash on a malformed
> rich-text value, and must localize the fault — is captured here; promote to a corpus case if a
> reliably-reproducing malformed shape is found.

## Desired behavior

- Serializing a workbook **never crashes with an unguarded undefined-property access** when a cell
  carries a malformed/incomplete rich-text-shaped value. The writer either:
  1. treats a rich-text value with a missing/empty runs array as an **empty/blank cell**, or
  2. **fails fast with a precise, actionable error** that names the offending **sheet + cell
     address** and states that the rich-text value is missing its runs array.
- Silent deep-stack `TypeError`s that do not identify the cell are unacceptable — the whole point is
  that a caller building values programmatically can find the one bad cell.
- A well-formed rich-text value (a `richText` array of styled runs) continues to serialize and
  round-trip unchanged (already exercised by the rich-text corpus cases).

## Open questions

- Default policy: tolerate-as-blank (robust, but hides a caller bug) vs fail-fast-with-location
  (surfaces the bug)? A strict/validate mode could choose fail-fast while the default tolerates.
- How broadly to validate cell-value shapes on write — just rich text, or a general "value shape"
  guard that localizes any malformed value (formula, hyperlink, date) to its cell address?
- Should the type surface make a malformed rich-text value unrepresentable (a precise `RichText`
  type) so the crash class is prevented at compile time for TypeScript callers?

Related: `hyperlink-display-text-can-be-rich-text`, `streaming-write-richtext-shared-strings-distinct`,
`html-fragment-to-rich-text-cell-value`, `public-type-surface-matches-runtime`.
