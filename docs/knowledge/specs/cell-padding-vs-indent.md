# Cell "padding": be honest about indent, don't invent unroundtrippable padding

Cluster: styles

## Scenario

Users formatting information-dense spreadsheets want breathing room inside cells — space between the
cell border and its text, like CSS padding in HTML tables — and ask for a "cell padding" setting. The
target format (OOXML) has no general four-sided per-cell padding, so a CSS-style padding API would be
a value that cannot round-trip: it would either be silently dropped on write or faked in a way that
does not survive reopening in Excel.

> Spec note, not a corpus case: this is an API-design and honesty question, not a failing behavior.
> The durable value is the stance — expose the format's real inset mechanism, don't invent one that
> lies about round-trip fidelity.

## Format reality

- OOXML cell alignment exposes an integer **`indent`** (0..~250) that offsets the text of left- or
  right-aligned cells by whole indent levels. This is the only native inset, and it is **one-sided**
  (leading edge only), not four-sided padding. The reference implementation already exposes
  `alignment.indent`.
- There is **no native top/bottom/right padding** in pixels or points. Vertical spacing comes only
  from row height; trailing-side horizontal spacing only from column width.
- Google Sheets / HTML export layers sometimes fake padding with indent or spaces; those do not
  survive as true padding in `.xlsx`.

## Desired stance

- **Do not introduce a CSS-style four-sided `padding`** on the cell model — it cannot round-trip and
  would mislead authors into expecting fidelity the format cannot provide.
- Expose the real mechanism clearly: `alignment.indent` for leading-edge inset, documented as
  one-sided and level-based, with row height / column width as the axes for the other directions.
- If an ergonomic helper is wanted, it should map explicitly onto indent + row height + column width
  and **document that it is an approximation**, not a faithful padding model.

## Open questions

- Is an approximating helper worth the confusion, or is documenting the native knobs (indent, row
  height, column width) sufficient?
- Should a set-`padding` attempt throw/warn (honest rejection) rather than silently degrade to indent?

Related: `set-style-over-cell-range`, `alignment-does-not-leak-across-cells`,
`worksheet-to-html-export`.
