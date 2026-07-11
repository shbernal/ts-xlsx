# A defined name may hold an arbitrary formula, not only a range

Cluster: formulas

## Scenario

A workbook's Name Manager entries (defined names) must be able to hold arbitrary formula
expressions, not only plain cell/range references. Excel stores every defined name as a formula
string — a range reference like `Sheet1!$A$2:$A$6` is merely the simplest such expression. The
standard technique for cascading / dynamic dropdowns is a defined name whose definition is a dynamic
range built with `INDEX`/`MATCH` or `OFFSET`, e.g.:

```
=INDEX(department_filter1_table, , MATCH(department, departments_list, 0))
```

referenced as the list source of a data validation, so the options offered by one dropdown depend on
the value chosen in another. The library currently accepts only a range address when adding a defined
name, so these formula-valued names cannot be authored at all.

> Spec note, not a corpus case: this is a feature gap (the defined-names surface accepts only ranges),
> not a bug with a reproduction, and the fork has no defined-names adapter capability yet. The durable
> value is the formula-first defined-name model and its wiring to the data-validation list source. It
> becomes a corpus case once the API accepts a formula-valued name and a round-trip asserts the
> emitted `<definedName>` text is preserved verbatim.

## Desired behavior

- **A defined name's definition is formula text.** The API accepts any legal formula string as the
  value of a name — `INDEX`/`MATCH`, `OFFSET`, calls over other named ranges — not only a bare range.
  A plain range is just the simplest accepted formula shape.
- **The definition is preserved verbatim on write** (correctly XML-escaped) and parsed back on read,
  without being validated or normalized as a range. An unrecognized-but-legal expression is not
  rejected.
- **Scope is preserved** whether the definition is a range or a general formula: a name may be
  workbook-scoped or sheet-scoped (`localSheetId`), and that scope round-trips unchanged. See
  `defined-name-scope-must-be-per-sheet`.
- **A data-validation list source may reference a defined name**, so a file authored with a named
  dynamic range as a dropdown's source opens in Excel with a working (cascading) dropdown. See
  `list-validation-inline-formula-length-limit`.

## Prior art (OOXML)

In `xl/workbook.xml`, `<definedNames><definedName name="…" localSheetId="…">EXPRESSION</definedName>`
stores the target as a formula expression. Excel's Name Manager treats every name this way; the
range-only restriction is a library limitation, not a format one.

## Open questions

- Public API shape: a single `formula`/`definition` string field on a defined name, with a plain
  range accepted as one formula shape — versus keeping a range-oriented convenience. Prefer
  formula-first.
- Whether to offer lightweight validation (balanced parens, known-name references) or store the
  expression opaquely and defer correctness to the host. Leaning opaque-but-preserved to avoid
  rejecting valid-but-unrecognized syntax.
- How absolute/relative references and cross-sheet qualification are normalized on round-trip, if at
  all — verbatim preservation is the safest baseline.

Related: `defined-name-scope-must-be-per-sheet`, `defined-names-tolerate-non-address-tokens`,
`list-validation-inline-formula-length-limit`, `formula-cell-value-type-minimal-required-fields`.
