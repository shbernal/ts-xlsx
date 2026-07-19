# Cell values

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellValue`

<sub>type</sub>

Everything a cell's value can be. `null` is the empty cell.

```ts
type CellValue = null | number | string | boolean | Date | ErrorValue | FormulaValue | SharedFormulaValue | DataTableFormulaValue | RichTextValue | HyperlinkValue;
```

---

### `coerceCellValue`

<sub>function</sub>

Normalise a raw assignment into a stored `CellValue`. `undefined` becomes the
empty cell (`null`); every other kind is validated by `detectValueType`. The
model never rewrites one value *kind* into another (a numeric-looking string stays a
string) — the single exception is formula text, which is canonicalised to the OOXML
stored form (no leading `=`) so round-trips are idempotent regardless of how the
caller supplied it.

```ts
function coerceCellValue(value: CellValue | undefined): CellValue;
```

**Throws** — if the value is not a recognised cell-value shape.

---

### `detectValueType`

<sub>function</sub>

Classify a value into its observable `ValueType`. This is total over
`CellValue`: every legal value has exactly one type. A `Date` is a date even
when its time is `NaN` (an invalid date is still a date-typed cell); serialization,
not the model, decides what to do with it.

```ts
function detectValueType(value: CellValue): ValueType;
```

---

### `ERROR_CODES`

<sub>const</sub>

The canonical Excel error literals a cell (or formula result) can carry.

```ts
const ERROR_CODES: readonly ["#N/A", "#REF!", "#NAME?", "#DIV/0!", "#NULL!", "#VALUE!", "#NUM!", "#SPILL!", "#CALC!", "#GETTING_DATA"]
```

---

### `ErrorCode`

<sub>type</sub>

```ts
type ErrorCode = (typeof ERROR_CODES)[number];
```

---

### `ErrorValue`

<sub>interface</sub>

An in-cell error, e.g. `{error: '#REF!'}`.

```ts
interface ErrorValue {
    readonly error: ErrorCode;
}
```

---

### `FormulaResult`

<sub>type</sub>

The cached result a formula carries — any scalar, a date, or an error.

```ts
type FormulaResult = number | string | boolean | Date | ErrorValue;
```

---

### `FormulaValue`

<sub>interface</sub>

A cell whose value is computed by its own formula.

```ts
interface FormulaValue {
    readonly formula: string;
    readonly result?: FormulaResult;
}
```

---

### `HyperlinkValue`

<sub>interface</sub>

A hyperlink cell: a URL plus the text (plain or rich) shown in the cell.

```ts
interface HyperlinkValue {
    readonly hyperlink: string;
    readonly text: string | RichTextValue;
    readonly tooltip?: string;
    readonly range?: string;
}
```

---

### `isErrorCode`

<sub>function</sub>

Whether a string is one of Excel's canonical error literals.

```ts
function isErrorCode(text: string): text is ErrorCode;
```

---

### `RichTextRun`

<sub>interface</sub>

One formatted run of a rich-text value.

```ts
interface RichTextRun {
    readonly text: string;
    readonly font?: Partial<Font>;
}
```

---

### `richTextToPlain`

<sub>function</sub>

Flatten a rich-text value to its plain text by concatenating every run's text in order. This is the
text a consumer that cannot render per-run formatting (a CSV field, a pivot cache entry) sees, and
the string a rich cell reads as when its formatting is discarded.

```ts
function richTextToPlain(value: RichTextValue): string;
```

---

### `RichTextValue`

<sub>interface</sub>

A value composed of independently-formatted text runs.

```ts
interface RichTextValue {
    readonly richText: readonly RichTextRun[];
}
```

---

### `SharedFormulaValue`

<sub>interface</sub>

A cell that participates in a shared formula — a clone of a master formula cell filled across a
range. `sharedFormula` is the master cell's address (e.g. `'B1'`); the master itself is a plain
`FormulaValue`. On read, the clone's own formula is the master's translated to the clone's
position and `result` is the clone's cached value; on write, the clones of a master collapse into
OOXML's shared-formula grouping.

```ts
interface SharedFormulaValue {
    readonly sharedFormula: string;
    readonly formula?: string;
    readonly result?: FormulaResult;
}
```

---

### `ValueType`

<sub>const</sub>

The observable kind of a cell's value. Both formula shapes report as `Formula`.

```ts
const ValueType: { readonly Null: "null"; readonly Number: "number"; readonly String: "string"; readonly Boolean: "boolean"; readonly Date: "date"; readonly Error: "error"; readonly Formula: "formula"; readonly RichText: "richText"; readonly Hyperlink: "hyperlink"; }
```
