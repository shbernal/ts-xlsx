// Conditional formatting: the rules that restyle a cell based on its value (a data bar, a colour
// scale, a "highlight cells greater than 10", a formula-driven expression). Like a data validation,
// it is a worksheet-level overlay keyed by a target range, not a facet owned by one cell: one rule
// covers a whole range, and several rules can layer on the same cells with an evaluation precedence.
//
// The model carries the operands each rule type needs and otherwise leaves them absent. A rule type
// the library does not interpret in depth still round-trips its `type`, `priority`, `operator`,
// `formulae`, and differential-style reference, so a read/write cycle never silently drops a rule.

import {tokenSet} from '../token-set.ts';
import type {Color, DifferentialStyle} from './style.ts';

/** How a {@link CfValueObject} reads its `value`: `ST_CfvoType` verbatim. */
export type CfValueObjectType = 'num' | 'percent' | 'max' | 'min' | 'percentile' | 'formula';

/** Narrow a raw `<cfvo type>` token to a known {@link CfValueObjectType}. */
export const isCfValueObjectType = tokenSet<CfValueObjectType>({
  num: true,
  percent: true,
  max: true,
  min: true,
  percentile: true,
  formula: true,
});

/**
 * One anchor of a colour-scale, data-bar, or icon-set scale: a "conditional format value object".
 * `type` names how `value` is read: a literal `num`, a `percent`/`percentile` of the range, a
 * `formula`, or the range's own `min`/`max` (which carry no value).
 */
export interface CfValueObject {
  type: CfValueObjectType;
  value?: number | string;
}

/**
 * What a rule tests, as `<cfRule type>` carries it: `ST_CfType` verbatim.
 *
 * Closed, and stated in full rather than left as `string`, even though the library models only some
 * of these in depth. Depth of modelling and legality are different questions: a `timePeriod` rule
 * whose operands the library never inspects still round-trips, while a token outside this list is one
 * Excel refuses to open, so it is refused on the way in and on the way out alike.
 */
export type ConditionalFormattingType =
  | 'expression'
  | 'cellIs'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'top10'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod'
  | 'aboveAverage';

/** Narrow a raw `<cfRule type>` token to a known {@link ConditionalFormattingType}. */
export const isConditionalFormattingType = tokenSet<ConditionalFormattingType>({
  expression: true,
  cellIs: true,
  colorScale: true,
  dataBar: true,
  iconSet: true,
  top10: true,
  uniqueValues: true,
  duplicateValues: true,
  containsText: true,
  notContainsText: true,
  beginsWith: true,
  endsWith: true,
  containsBlanks: true,
  notContainsBlanks: true,
  containsErrors: true,
  notContainsErrors: true,
  timePeriod: true,
  aboveAverage: true,
});

/**
 * How a `cellIs` or text rule compares, as `ST_ConditionalFormattingOperator` enumerates it.
 *
 * Overlaps {@link import('./data-validation.ts').DataValidationOperator} in four members and diverges
 * in the rest: this one has the text comparisons a validation has no use for, and spells "does not
 * contain" as `notContains` where nothing else in the format does.
 */
export type ConditionalFormattingOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween'
  | 'containsText'
  | 'notContains'
  | 'beginsWith'
  | 'endsWith';

/** Narrow a raw `<cfRule operator>` token to a known {@link ConditionalFormattingOperator}. */
export const isConditionalFormattingOperator = tokenSet<ConditionalFormattingOperator>({
  lessThan: true,
  lessThanOrEqual: true,
  equal: true,
  notEqual: true,
  greaterThanOrEqual: true,
  greaterThan: true,
  between: true,
  notBetween: true,
  containsText: true,
  notContains: true,
  beginsWith: true,
  endsWith: true,
});

/** The window a `timePeriod` rule matches against, relative to the day the sheet is recalculated. */
export type CfTimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

/** Narrow a raw `<cfRule timePeriod>` token to a known {@link CfTimePeriod}. */
export const isCfTimePeriod = tokenSet<CfTimePeriod>({
  today: true,
  yesterday: true,
  tomorrow: true,
  last7Days: true,
  thisMonth: true,
  lastMonth: true,
  nextMonth: true,
  thisWeek: true,
  lastWeek: true,
  nextWeek: true,
});

/**
 * The named icon family an `iconSet` rule draws from, as `ST_IconSetType` enumerates it. The leading
 * digit is the number of icons, which is also how many {@link CfValueObject} anchors the rule needs.
 *
 * The 2009 extension adds three more families (`3Stars`, `3Triangles`, `5Boxes`) under its own
 * namespace. They are absent here because the classic `<iconSet>` element this list types cannot
 * carry them; a file using one states it in the extension, which the library round-trips verbatim.
 */
export type IconSetType =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters';

/** Narrow a raw `<iconSet iconSet>` token to a known {@link IconSetType}. */
export const isIconSetType = tokenSet<IconSetType>({
  '3Arrows': true,
  '3ArrowsGray': true,
  '3Flags': true,
  '3TrafficLights1': true,
  '3TrafficLights2': true,
  '3Signs': true,
  '3Symbols': true,
  '3Symbols2': true,
  '4Arrows': true,
  '4ArrowsGray': true,
  '4RedToBlack': true,
  '4Rating': true,
  '4TrafficLights': true,
  '5Arrows': true,
  '5ArrowsGray': true,
  '5Rating': true,
  '5Quarters': true,
});

/**
 * A single conditional-formatting rule. `type` is the OOXML cfRule type; the remaining fields carry
 * the operands that type needs and are absent otherwise. A rule the library does not model in depth
 * still preserves `type`, `priority`, `operator`, `formulae`, and `dxfId` across a round-trip.
 */
export interface ConditionalFormattingRule {
  type: ConditionalFormattingType;
  /** Evaluation precedence; lower wins. Excel requires one, so the writer supplies it when absent. */
  priority?: number;
  /** Halt evaluation of lower-priority rules on any cell this rule matches. */
  stopIfTrue?: boolean;
  /** cellIs / text comparison operator (`greaterThan`, `between`, `beginsWith`, …). */
  operator?: ConditionalFormattingOperator;
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
  iconSet?: IconSetType;
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
  timePeriod?: CfTimePeriod;
}

/** A set of rules bound to the range(s) they cover. `ref` is an OOXML `sqref`: one or more
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
