# Template placeholder replacement: fill a design template from a data context

Cluster: templating

## Scenario

A user starts from a pre-designed spreadsheet template that carries placeholder tokens — a text
marker like `{{customerName}}`, or an image placeholder — repeated across multiple sheets and
multiple cells. They want to load the template, supply a set of named values and images, and get back
a workbook where every occurrence of each token is replaced with the corresponding data: text with
text, image placeholders with real embedded images. Beyond flat substitution they also want
lightweight logic — conditional inclusion (`if`/`else`) and repetition over a collection (`foreach`
expanding rows/blocks driven by array data).

> Spec note, not a corpus case: this is a feature proposal with no reproduction and no single
> assertable behavior. The durable value is the capability's scope and the design decisions. If
> prototyped it composes existing primitives (cell read/mutate, `spliceRows` for `foreach` expansion,
> image insertion) rather than needing new low-level format support, so it is a layer above the
> document model — worth naming as a distinct, optional surface.

## Desired behavior

- **Text substitution.** A delimited token (mustache-style `{{name}}`) found in a cell's text/value is
  replaced by the bound value. Tokens resolve in every sheet and every cell they appear in. A token
  embedded within a larger string replaces in place (the surrounding text is kept); the cell's style
  and number format are preserved across the substitution.
- **Image substitution.** An image placeholder (a token, or a marked cell/anchor) is replaced by an
  actual embedded image supplied by name, anchored where the placeholder was, reusing the existing
  image-anchoring model.
- **Repetition (`foreach`).** A block or row region bound to a collection is expanded once per item,
  shifting subsequent rows down (the `spliceRows`-style insert), with per-item tokens resolving
  against each element. Merges, styles, and row heights in the template block are carried to each
  expansion.
- **Conditional inclusion (`if`/`else`).** A region guarded by a condition is kept or dropped (its
  rows removed) based on the data context.
- **Missing-binding policy is explicit.** A token with no bound value resolves to a defined outcome
  (empty string, left-as-is, or error) — chosen and documented, not silently divergent.

## Open questions

- Is templating in the library's core, or a separate opt-in package built on the document API? A
  separate layer keeps the core a pure reader/writer and matches the "small, composable" stance.
- Token syntax and escaping: mustache `{{ }}` by default; how does a caller emit a literal `{{`? How
  are nested paths (`{{customer.name}}`) and formatting directives handled?
- Expansion semantics: how do `foreach`/`if` interact with tables, merged regions, charts (which
  reference cell ranges), and defined names when rows are inserted/removed?
- Type fidelity: a substituted value that is a number/date must become a typed cell, not a string, so
  downstream formulas and formats work.

Related: `streaming-read-modify-write-template`, `chart-parts-survive-template-roundtrip`,
`in-cell-rich-value-images`, `worksheet-columns-mutable-array-ergonomics`.
