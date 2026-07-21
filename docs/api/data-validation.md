# Data Validation

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DataValidation`

<sub>interface</sub>

One validation rule. `formulae` holds the operand(s) — `formula1` then optional `formula2`: a
numeric literal is stored as a number, while a cell reference, defined name, or list source keeps
its verbatim string.

```ts
interface DataValidation {
    type: DataValidationType;
    operator?: DataValidationOperator;
    formulae?: (string | number)[];
    allowBlank?: boolean;
    showInputMessage?: boolean;
    showErrorMessage?: boolean;
    errorStyle?: DataValidationErrorStyle;
    error?: string;
    errorTitle?: string;
    prompt?: string;
    promptTitle?: string;
}
```

---

### `DataValidationEntry`

<sub>interface</sub>

A validation bound to the range(s) it covers. `sqref` is an OOXML `sqref` — one or more
space-separated ranges. `extended` marks a rule stored in the 2009 extension form
(`<x14:dataValidation>` inside the worksheet `<extLst>`) — Excel's carrier for validations a
legacy `<dataValidation>` cannot express, such as a list source on another sheet. The flag is how
a rule read from that form remembers to be written back to it, rather than downgraded to the
standard element (which would corrupt a cross-sheet reference).

```ts
interface DataValidationEntry {
    sqref: string;
    rule: DataValidation;
    extended?: boolean;
}
```

---

### `DataValidationErrorStyle`

<sub>type</sub>

How Excel reacts to input that fails the rule.

```ts
type DataValidationErrorStyle = 'stop' | 'warning' | 'information';
```

---

### `DataValidationOperator`

<sub>type</sub>

How a typed validation compares its operand(s). Absent on a `list`/`custom` rule; defaults to
`between` on a typed rule (the value Excel omits from the XML).

```ts
type DataValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
```

---

### `DataValidationType`

<sub>type</sub>

The kind of constraint a validation enforces. `list` is a dropdown; `custom` is an arbitrary
boolean formula; the rest bound a typed value (`whole`/`decimal`/`date`/`time`/`textLength`).

```ts
type DataValidationType = 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
```
