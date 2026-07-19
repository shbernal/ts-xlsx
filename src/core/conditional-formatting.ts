// Conditional formatting — the rules that restyle a cell based on its value (a data bar, a colour
// scale, a "highlight cells greater than 10", a formula-driven expression). Like a data validation,
// it is a worksheet-level overlay keyed by a target range, not a facet owned by one cell: one rule
// covers a whole range, and several rules can layer on the same cells with an evaluation precedence.
//
// The model carries the operands each rule type needs and otherwise leaves them absent. A rule type
// the library does not interpret in depth still round-trips its `type`, `priority`, `operator`,
// `formulae`, and differential-style reference — so a read/write cycle never silently drops a rule.

import type {Border, Color, Fill, Font} from './style.ts';

/**
 * A differential style (OOXML CT_Dxf): the facets a matching rule lays over a cell's own formatting.
 * Only the facets present override; the rest of the cell's style shows through. This is the format a
 * `highlight`/`cellIs`/`expression` rule applies, distinct from the colour/anchor scale a
 * `dataBar`/`colorScale`/`iconSet` rule renders.
 */
export interface DifferentialStyle {
  font?: Partial<Font>;
  numFmt?: string;
  fill?: Fill;
  border?: Border;
}

/**
 * One anchor of a colour-scale, data-bar, or icon-set scale — a "conditional format value object".
 * `type` names how `value` is read: a literal `num`, a `percent`/`percentile` of the range, a
 * `formula`, or the range's own `min`/`max` (which carry no value).
 */
export interface CfValueObject {
  type: 'num' | 'percent' | 'max' | 'min' | 'percentile' | 'formula';
  value?: number | string;
}

/**
 * A single conditional-formatting rule. `type` is the OOXML cfRule type; the remaining fields carry
 * the operands that type needs and are absent otherwise. A rule the library does not model in depth
 * still preserves `type`, `priority`, `operator`, `formulae`, and `dxfId` across a round-trip.
 */
export interface ConditionalFormattingRule {
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

/** A set of rules bound to the range(s) they cover. `ref` is an OOXML `sqref` — one or more
 * space-separated areas (`"A1:C1 A3:C3 A5:C5"`), the shape Excel writes when one rule is applied to
 * several non-contiguous selections at once. */
export interface ConditionalFormatting {
  ref: string;
  rules: ConditionalFormattingRule[];
}

/** A defensive deep copy, so a stored conditional formatting never aliases the caller's object nor
 * any of its nested arrays (rules, formulae, cfvo, colours) or the differential style. */
export function cloneConditionalFormatting(cf: ConditionalFormatting): ConditionalFormatting {
  return {ref: cf.ref, rules: cf.rules.map(cloneRule)};
}

function cloneRule(rule: ConditionalFormattingRule): ConditionalFormattingRule {
  return {
    ...rule,
    ...(rule.formulae !== undefined ? {formulae: [...rule.formulae]} : {}),
    ...(rule.cfvo !== undefined ? {cfvo: rule.cfvo.map((v) => ({...v}))} : {}),
    ...(rule.color !== undefined ? {color: {...rule.color}} : {}),
    ...(rule.negativeFillColor !== undefined
      ? {negativeFillColor: {...rule.negativeFillColor}}
      : {}),
    ...(rule.axisColor !== undefined ? {axisColor: {...rule.axisColor}} : {}),
    ...(rule.colors !== undefined ? {colors: rule.colors.map((c) => ({...c}))} : {}),
    ...(rule.style !== undefined ? {style: cloneStyle(rule.style)} : {}),
  };
}

function cloneStyle(style: DifferentialStyle): DifferentialStyle {
  return {
    ...style,
    ...(style.font !== undefined ? {font: {...style.font}} : {}),
    ...(style.fill !== undefined ? {fill: {...style.fill}} : {}),
    ...(style.border !== undefined ? {border: {...style.border}} : {}),
  };
}
