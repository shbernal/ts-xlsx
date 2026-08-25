# Author a rich-text cell value from an inline-styled markup fragment

Cluster: styles

## Scenario

Users assembling reports often already hold small fragments of inline-styled markup (bold, italic,
colored spans, line breaks, simple paragraphs) and want to drop that into a single cell so it renders
as formatted, mixed-run text. Spreadsheets support this natively via rich text, where a cell value can
be a sequence of runs, each with its own font properties. Today authors must hand-translate their
markup into an explicit array of run objects, which is tedious and error-prone. The request is a
convenience path: assign a markup fragment to a cell and have the library parse the inline styling
into the equivalent rich-text run sequence.

> Spec note, not a corpus case: this is a convenience feature with no failing behavior to assert,
> since the underlying rich-text run model already works and round-trips. The durable value is the
> mapping contract from markup onto that model, plus the security posture, because it parses untrusted
> input.

## Desired behavior

- An ergonomic way to build a rich-text cell value from an inline-styled markup fragment, translating
  supported formatting into the native run model rather than storing literal markup as plain text.
- **Inline tags map to run font properties**: `<b>`/`<strong>` to bold, `<i>`/`<em>` to italic, `<u>`
  to underline, `<s>`/`<del>` to strike, `<span style="color:…">` to color, with size and family from
  `style`.
- **Structural whitespace produces newlines within the one cell**: `<br>` gives a line break, and
  block boundaries (`</p>`, block-level elements) give paragraph separation. A multi-line cell implies
  wrap-text alignment so the lines are visible.
- **Nested and overlapping formatting composes**: bold text containing a colored span yields runs
  carrying both properties on the overlap.
- **The output is a normal rich-text value**: reading the cell back yields the run sequence, and a
  package round-trip preserves the runs and their font properties.
- **Unknown tags and attributes are handled deterministically**, stripped to their text content rather
  than silently dropping the text, and the parser is hardened against hostile input, with no unbounded
  expansion and no script execution. This is untrusted-input-facing.

## Open questions

- Which markup dialect: a curated inline subset, which is safer and smaller to secure, or a broader
  HTML parse? A narrow documented subset is preferred.
- CSS coverage: how much of `style="…"` (named against hex against rgb colors, numeric font-weight,
  text-decoration combinations) is in scope.
- API shape: a distinct value kind, or a helper that returns a rich-text value the caller assigns? A
  helper returning the native rich-text value is more composable and leaves the cell value model
  unchanged.
- Is the inverse (rich text to HTML export) also wanted? Several users conflate the two directions.
  This is not a full HTML/CSS layout engine and should not render tables, images or the box model in a
  cell.

Related: `formula-string-result-under-date-format-roundtrip`, `set-style-over-cell-range`.
