# Column-level value type

Cluster: types

Reported as "set columns type".

## Desired behavior

A column definition should support declaring the intended value type or semantic for the whole column, so authors don't have to set formatting and coercion cell-by-cell. Setting this on a column should:

1. Apply a sensible default number format to every cell in the column, so a text column maps to the `@` format and values are preserved as-entered rather than reinterpreted by the spreadsheet app.
2. Optionally influence how raw values written into the column are coerced, so a text column keeps `"007"` as the string `"007"` rather than the number `7`, and a number column parses numeric-looking strings to numbers.
3. Be inheritable by cells created later in that column, so rows added after the column is configured still pick up the column's type.

## Prior art

- The upstream library exposed a `ValueType` enum and per-cell type inference, but never honored a `type` field on column definitions, and supplying one was silently dropped. This request accumulated many duplicate "+1"s over years, indicating real demand.
- The community workaround was to set the column's number format to text (`numFmt = '@'`). This is display-only: it does not stop value coercion, and it is not discoverable from the column API.
- Excel and OOXML have no first-class "column data type" concept; column typing is a library-level convenience layered on top of per-cell number formats and cell value types.

## Open questions

- Should the field control only default formatting, only value coercion, or both? Recommendation: both, since that is what the workaround fails to deliver.
- What is the field's name and value space? Reusing the existing value-type enum is intuitive, but that enum mixes concerns (formula, hyperlink, rich text) that don't make sense as a column-wide coercion target. A narrower column-type set (text, number, date, boolean) may be clearer.
- Precedence: when both a column type and an explicit per-cell number format are set, the per-cell setting should win.
- Text columns and leading zeros or long numeric IDs are the highest-value case and should be the primary acceptance scenario.
