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
  /** Evaluation precedence; lower wins. Excel requires one — the writer supplies it when absent. */
  priority?: number;
  /** Halt evaluation of lower-priority rules on any cell this rule matches. */
  stopIfTrue?: boolean;
  /** cellIs / text comparison operator (`greaterThan`, `between`, `beginsWith`, …). */
  operator?: string;
  /** Formula operands: cellIs bounds, an expression predicate, a containsText target formula, … */
  formulae?: (string | number)[];
  /** The literal a containsText / beginsWith / endsWith rule searches for. */
  text?: string;
  /** A differential style authored inline, serialised into `<dxfs>` and referenced by the cfRule. */
  style?: DifferentialStyle;
  /** A differential-style reference by `<dxfs>` index, as read from a file (kept verbatim). */
  dxfId?: string;
  /** colorScale / dataBar / iconSet scale anchors, in order. */
  cfvo?: CfValueObject[];
  /** A dataBar's bar colour. */
  color?: Color;
  /** A colorScale's colours, one per {@link cfvo}. */
  colors?: Color[];
  /** A dataBar's gradient-fill flag. Lives only in the x14 extension, not the classic element. */
  gradient?: boolean;
  /** A dataBar's fill colour for negative values. An x14 extension property. */
  negativeFillColor?: Color;
  /** A dataBar's axis colour (the zero line between positive and negative bars). An x14 property. */
  axisColor?: Color;
  /** An iconSet's named icon family (e.g. `3TrafficLights1`). */
  iconSet?: string;
  /** top10 rank cutoff. */
  rank?: number;
  /** top10: the rank is a percentage rather than a count. */
  percent?: boolean;
  /** top10: rank from the bottom rather than the top. */
  bottom?: boolean;
  /** aboveAverage: match above (default) or below the average. */
  aboveAverage?: boolean;
  /** aboveAverage: include cells equal to the average. */
  equalAverage?: boolean;
  /** aboveAverage: match beyond this many standard deviations. */
  stdDev?: number;
  /** timePeriod window (`today`, `lastWeek`, …). */
  timePeriod?: string;
}
```
