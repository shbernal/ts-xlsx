# Autofilter

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AutoFilter`

<sub>interface</sub>

A worksheet's autofilter: the filtered region plus any per-column criteria narrowing it. A bare
range (no columns) is just the header-row dropdowns Excel draws; adding `FilterColumn`s
records the criteria a column is actively filtered by.

```ts
interface AutoFilter {
    readonly ref: string;
    readonly columns: readonly FilterColumn[];
}
```

---

### `CustomFilter`

<sub>interface</sub>

A column narrowed to one or two operator predicates (`> 6`, `<> "draft"`). Two predicates are
AND-combined when `and` is set, else OR-combined; Excel permits at most two.

```ts
interface CustomFilter {
    readonly kind: 'custom';
    readonly and: boolean;
    readonly predicates: readonly CustomFilterPredicate[];
}
```

---

### `CustomFilterOperator`

<sub>type</sub>

```ts
type CustomFilterOperator = 'equal' | 'notEqual' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';
```

---

### `CustomFilterPredicate`

<sub>interface</sub>

```ts
interface CustomFilterPredicate {
    readonly operator: CustomFilterOperator;
    readonly val: string;
}
```

---

### `FilterColumn`

<sub>interface</sub>

One filtered column, addressed by its 0-based offset (`colId`) from the filter range's left edge.

```ts
interface FilterColumn {
    readonly colId: number;
    readonly criteria: FilterCriteria;
}
```

---

### `FilterCriteria`

<sub>type</sub>

The two criteria kinds this library models: a discrete value set, or operator predicates.

```ts
type FilterCriteria = ValuesFilter | CustomFilter;
```

---

### `ValuesFilter`

<sub>interface</sub>

A column narrowed to a discrete set of allowed values — the checkbox list in Excel's dropdown.
A row survives when its cell in this column matches one of `values` (or is blank, when
`blank` is set).

```ts
interface ValuesFilter {
    readonly kind: 'values';
    readonly values: readonly string[];
    readonly blank: boolean;
}
```
