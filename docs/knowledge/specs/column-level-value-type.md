# set columns type

## Column-level value type

### Desired behavior
A column definition should support declaring the intended value type / semantic for the whole column, so authors don't have to set formatting and coercion cell-by-cell. Setting this on a column should:

1. Apply a sensible default number format to every cell in the column (e.g. a text column maps to the `@` format so values are preserved as-entered and not reinterpreted by the spreadsheet app).
2. Optionally influence how raw values written into the column are coerced (e.g. a text column keeps `"007"` as the string `"007"` rather than the number `7`; a number column parses numeric-looking strings to numbers).
3. Be inheritable by cells created later in that column (rows added after the column is configured still pick up the column's type).

### Prior art
- The upstream library exposed a `ValueType` enum and per-cell type inference, but never honored a `type` field on column definitions — supplying one was silently dropped. This request accumulated many duplicate "+1"s over years, indicating real demand.
- The community workaround was to set the column's number format to text (`numFmt = '@'`). This is display-only: it does not stop value coercion, and it is not discoverable from the column API.
- Excel/OOXML has no first-class "column data type" concept; column typing is a library-level convenience layered on top of per-cell number formats and cell value types.

### Open questions
- Should the field control only default formatting, only value coercion, or both? (Recommendation: both, since that is what the workaround fails to deliver.)
- What is the field's name and value space? Reusing the existing value-type enum is intuitive, but that enum mixes concerns (formula, hyperlink, rich text) that don't make sense as a column-wide coercion target. A narrower column-type set (text / number / date / boolean) may be clearer.
- Precedence: when both a column type and an explicit per-cell number format are set, the per-cell setting should win.
- Text columns and leading zeros / long numeric IDs are the highest-value case and should be the primary acceptance scenario.
