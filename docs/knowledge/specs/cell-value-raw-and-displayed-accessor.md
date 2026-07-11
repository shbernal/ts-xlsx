# A unified accessor for a cell's raw and displayed value

Cluster: types

## Scenario

A consumer iterating cells wants "the value" without branching over every value-type variant. Today a
cell's `value` is a discriminated union — a number, a string, a `Date`, a rich-text object, a
hyperlink object `{text, hyperlink}`, a formula object `{formula, result}`, an error, a boolean — so a
caller who just wants the underlying scalar, or just the string a spreadsheet would display, must
hand-write a switch over all of them. The recurring ask is for one accessor that returns the *raw*
value (the underlying scalar, formula result unwrapped, hyperlink text extracted) and/or the
*displayed* value (the string the application renders, number-format applied).

> Spec note, not a corpus case: this is an API-ergonomics design decision with no failing current
> behavior to baseline — the union is already exposed and correct, the gap is a convenience surface on
> top of it. Recording the shape and the open questions feeds Phase 3 design.

## Desired behavior

Offer two distinct, clearly-named accessors so a caller never has to destructure the union by hand:

- **Raw value.** The underlying scalar, with wrapper objects unwrapped: a formula cell yields its
  cached `result`, a hyperlink cell yields its display text (or a structured `{text, target}` if the
  caller wants the link), rich text collapses to its concatenated plain text, a date stays a `Date`, a
  number stays a number, an error surfaces as a typed error value. This is what a caller means by "just
  give me the data" for export/serialization.

- **Displayed value.** The string the spreadsheet application would render for the cell: the raw value
  with the cell's effective number format applied (so `0.5` under a percent format reads `"50%"`, a
  date serial under a date format reads the formatted date, etc.). This requires a number-format
  formatter and must honor the workbook's locale/`date1904` epoch, consistent with the read-time date
  policy (see `xlsx-date-detection-control`).

- **Kind inspection.** Pair the accessors with a way to ask a value's kind (number / string / date /
  formula / hyperlink / rich-text / boolean / error) without brittle `typeof`/`instanceof` probing, so
  a caller can validate a column's contents robustly. This is the same inspectable-kind affordance the
  date-detection note calls for, generalized to every value type.

The type surface must make both accessors precisely typed: the raw accessor's return is the unwrapped
union; the displayed accessor returns `string` (with a defined result for empty/null cells).

## Open questions

- Naming: `cell.text` already exists in prior art as a display-ish string — does it become the
  "displayed value" accessor, or do we introduce distinct `raw`/`display` members to avoid overloading
  a name whose semantics were fuzzy?
- Does the displayed accessor build in a full number-format formatter (locale-aware), or start with a
  documented subset and defer exotic format codes?
- Hyperlink and rich-text cells: does "raw" flatten to a plain scalar/string by default, with the
  structured form available on request, or the reverse?
- Should these be lazy accessors computed on read, or precomputed — given the formatter cost on large
  sheets?

Related: `xlsx-date-detection-control`, `column-level-value-type`,
`formula-cell-value-type-minimal-required-fields`, `html-fragment-to-rich-text-cell-value`,
`public-type-surface-matches-runtime`.
