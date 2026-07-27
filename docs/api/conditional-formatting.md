# Conditional Formatting

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CfValueObject`

<sub>interface</sub>

One anchor of a colour-scale, data-bar, or icon-set scale — a "conditional format value object".
`type` names how `value` is read: a literal `num`, a `percent`/`percentile` of the range, a
`formula`, or the range's own `min`/`max` (which carry no value).

```ts
interface CfValueObject {
    type: 'num' | 'percent' | 'max' | 'min' | 'percentile' | 'formula';
    value?: number | string;
}
```

---

### `ConditionalFormatting`

<sub>interface</sub>

A set of rules bound to the range(s) they cover. `ref` is an OOXML `sqref` — one or more
space-separated areas (`"A1:C1 A3:C3 A5:C5"`), the shape Excel writes when one rule is applied to
several non-contiguous selections at once.

```ts
interface ConditionalFormatting {
    ref: string;
    rules: ConditionalFormattingRule[];
}
```

---

### `ConditionalFormattingRule`

<sub>interface</sub>

A single conditional-formatting rule. `type` is the OOXML cfRule type; the remaining fields carry
the operands that type needs and are absent otherwise. A rule the library does not model in depth
still preserves `type`, `priority`, `operator`, `formulae`, and `dxfId` across a round-trip.

```ts
interface ConditionalFormattingRule {
    type: string;
    priority?: number;
    stopIfTrue?: boolean;
    operator?: string;
    formulae?: (string | number)[];
    text?: string;
    style?: DifferentialStyle;
    dxfId?: string;
    cfvo?: CfValueObject[];
    color?: Color;
    colors?: Color[];
    gradient?: boolean;
    negativeFillColor?: Color;
    axisColor?: Color;
    iconSet?: string;
    rank?: number;
    percent?: boolean;
    bottom?: boolean;
    aboveAverage?: boolean;
    equalAverage?: boolean;
    stdDev?: number;
    timePeriod?: string;
}
```
