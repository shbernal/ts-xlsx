# Internal (in-workbook) hyperlinks must be portable across consumers

Cluster: xlsx-io

## Scenario

An author creates a hyperlink whose destination is another location inside the same workbook — a
cell reference like `Sheet2!A1`, or a defined name — rather than an external URL. In desktop Excel
the link navigates correctly. But when the same file is opened in Google Sheets or WPS Office the
link is dead: those consumers do not resolve it. The cause is how the internal link is serialized —
if it is emitted like an external link (an `r:id` relationship) instead of with an in-workbook
`location`, stricter consumers silently ignore it.

> Spec note, not a corpus case: the correct serialization is knowable, but the *observable* defect is
> cross-application navigation (a rendering behavior in third-party apps) that the corpus cannot
> exercise. The durable value is the OOXML shape an internal link must take.

## Desired behavior

- An internal hyperlink (to a cell/range on a sheet, or a defined name) is written so **non-Excel
  consumers** (Google Sheets, WPS) resolve and navigate to the target — not only Microsoft Excel.
- Concretely: an internal link carries a `location` attribute holding the in-workbook reference
  (e.g. `Sheet2!A1` or a defined name) and does **not** rely on an external-mode `r:id` relationship.
  An external link keeps its `r:id` → `.rels` target with `TargetMode="External"`; the two forms are
  serialized distinctly.
- A round-trip preserves the internal target verbatim (sheet-qualified reference or defined name),
  and the display text and tooltip survive.

## Open questions

- Does the writer infer "internal vs external" from the target shape (a sheet-qualified reference or
  a known defined name → internal; a URL scheme → external), or from an explicit caller flag?
- How are internal links to a defined name vs a raw cell reference distinguished, and are both
  emitted as `location`?
- Should an internal target that does not resolve to an existing sheet/name warn on write?

Related: `hyperlink-cell-default-style`, `cross-sheet-reference-preserved-in-formula-and-validation`,
`defined-name-full-row-column-span-survives-roundtrip`.
