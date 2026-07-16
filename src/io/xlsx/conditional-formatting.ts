// Serialising conditional formatting to the classic `<conditionalFormatting>` worksheet element.
//
// Each block names its target range(s) in a `sqref` attribute and holds one or more `<cfRule>`
// children. A rule's shape depends on its type: a `dataBar`/`colorScale`/`iconSet` carries a scale
// element (its `<cfvo>` anchors and colours), while a `cellIs`/`expression`/`top10`/… carries its
// operands as `<formula>` children and points at a differential style by `dxfId`. A rule the library
// does not model in depth still round-trips its attributes, so nothing is silently dropped on save.

import type {
  CfValueObject,
  ConditionalFormatting,
  ConditionalFormattingRule,
} from '../../core/conditional-formatting.ts';
import type {Color} from '../../core/style.ts';
import {colorAttrs, type StyleRegistry} from './styles.ts';
import {localName, parseXml} from './xml-read.ts';
import {escapeAttr, escapeText} from './xml.ts';

// Excel's default data bar when the author supplies none: a min/max anchor pair and its standard blue.
const DEFAULT_DATABAR_CFVO: readonly CfValueObject[] = [{type: 'min'}, {type: 'max'}];
const DEFAULT_DATABAR_COLOR: Color = {argb: 'FF638EC6'};

// The scale rules render a built-in visual and carry no differential style; the rest apply a dxf.
const SCALE_TYPES = new Set(['dataBar', 'colorScale', 'iconSet']);

/**
 * Serialise every conditional formatting on a sheet into its `<conditionalFormatting>` blocks, in
 * insertion order. Returns '' when the sheet has none. A rule missing a `priority` is assigned the
 * next free one, so the output always satisfies Excel's requirement that every cfRule carry one.
 */
export function conditionalFormattingsXml(
  formattings: readonly ConditionalFormatting[],
  styles: StyleRegistry
): string {
  if (formattings.length === 0) return '';
  const priority = {next: 1};
  return formattings.map(cf => blockXml(cf, styles, priority)).join('');
}

function blockXml(
  cf: ConditionalFormatting,
  styles: StyleRegistry,
  priority: {next: number}
): string {
  const rules = cf.rules.map(rule => ruleXml(rule, styles, priority)).join('');
  return `<conditionalFormatting sqref="${escapeAttr(cf.ref)}">${rules}</conditionalFormatting>`;
}

function ruleXml(
  rule: ConditionalFormattingRule,
  styles: StyleRegistry,
  priority: {next: number}
): string {
  const p = rule.priority ?? priority.next;
  // Keep the running counter ahead of any explicit priority so later auto-assigned ones stay unique.
  priority.next = Math.max(priority.next, p) + 1;

  const attrs = [`type="${escapeAttr(rule.type)}"`];
  const dxfId = resolveDxfId(rule, styles);
  if (dxfId !== undefined) attrs.push(`dxfId="${dxfId}"`);
  attrs.push(`priority="${p}"`);
  if (rule.stopIfTrue) attrs.push('stopIfTrue="1"');
  if (rule.aboveAverage === false) attrs.push('aboveAverage="0"');
  if (rule.equalAverage) attrs.push('equalAverage="1"');
  if (rule.bottom) attrs.push('bottom="1"');
  if (rule.percent) attrs.push('percent="1"');
  if (rule.operator !== undefined) attrs.push(`operator="${escapeAttr(rule.operator)}"`);
  if (rule.text !== undefined) attrs.push(`text="${escapeAttr(rule.text)}"`);
  if (rule.timePeriod !== undefined) attrs.push(`timePeriod="${escapeAttr(rule.timePeriod)}"`);
  if (rule.rank !== undefined) attrs.push(`rank="${rule.rank}"`);
  if (rule.stdDev !== undefined) attrs.push(`stdDev="${rule.stdDev}"`);

  const body = SCALE_TYPES.has(rule.type) ? scaleXml(rule) : formulaeXml(rule.formulae);
  return body === '' ? `<cfRule ${attrs.join(' ')}/>` : `<cfRule ${attrs.join(' ')}>${body}</cfRule>`;
}

// A rule points at a differential style either by a preserved index read from a file (`dxfId`) or by
// a style authored on the rule (interned here). The preserved index wins — it references the original
// file's dxf table, which the writer re-emits verbatim. Scale rules never carry one.
function resolveDxfId(rule: ConditionalFormattingRule, styles: StyleRegistry): number | undefined {
  if (SCALE_TYPES.has(rule.type)) return undefined;
  if (rule.dxfId !== undefined) return Number(rule.dxfId);
  if (rule.style !== undefined) return styles.differentialStyleId(rule.style);
  return undefined;
}

function formulaeXml(formulae: readonly (string | number)[] | undefined): string {
  if (formulae === undefined) return '';
  return formulae.map(f => `<formula>${escapeText(formulaText(f))}</formula>`).join('');
}

// A cellIs bound or expression predicate: a numeric literal serialises as its number, a string keeps
// its verbatim text with any single leading `=` stripped (authors write `=A1>0`, the element wants `A1>0`).
function formulaText(formula: string | number): string {
  if (typeof formula === 'number') return String(formula);
  return formula.startsWith('=') ? formula.slice(1) : formula;
}

function scaleXml(rule: ConditionalFormattingRule): string {
  if (rule.type === 'dataBar') return dataBarXml(rule);
  if (rule.type === 'colorScale') return colorScaleXml(rule);
  return iconSetXml(rule);
}

// A data bar states its low and high anchors and its bar colour. The minimal call (no cfvo, no colour)
// gains Excel's own defaults — a min/max anchor pair and the standard blue — rather than an invalid
// empty element. The gradient flag has no home in the classic element (it is an x14 extension), so it
// is not serialised here.
function dataBarXml(rule: ConditionalFormattingRule): string {
  const cfvo = rule.cfvo && rule.cfvo.length > 0 ? rule.cfvo : DEFAULT_DATABAR_CFVO;
  const color = rule.color ?? DEFAULT_DATABAR_COLOR;
  const anchors = cfvo.map(cfvoXml).join('');
  return `<dataBar>${anchors}<color ${colorAttrs(color)}/></dataBar>`;
}

// A colour scale pairs each anchor with a colour; a missing colour list falls back to none, still a
// well-formed (if plain) element.
function colorScaleXml(rule: ConditionalFormattingRule): string {
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  const colors = (rule.colors ?? []).map(c => `<color ${colorAttrs(c)}/>`).join('');
  return `<colorScale>${anchors}${colors}</colorScale>`;
}

function iconSetXml(rule: ConditionalFormattingRule): string {
  const name = rule.iconSet !== undefined ? ` iconSet="${escapeAttr(rule.iconSet)}"` : '';
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  return `<iconSet${name}>${anchors}</iconSet>`;
}

// One scale anchor. `min`/`max` carry no value; the rest state theirs in `val` (a formula anchor's
// value is its formula text, escaped like any attribute).
function cfvoXml(cfvo: CfValueObject): string {
  const val = cfvo.value !== undefined ? ` val="${escapeAttr(String(cfvo.value))}"` : '';
  return `<cfvo type="${escapeAttr(cfvo.type)}"${val}/>`;
}

// The scale kinds nest a `<color>` differently: a data bar names one bar colour, a colour scale a
// colour per anchor. Tracking which element we are inside routes a parsed `<color>` to the right slot.
type ScaleKind = 'dataBar' | 'colorScale' | 'iconSet';

// A rule under construction: fields accumulate across the cfRule's attributes and children, then are
// finalised into a ConditionalFormattingRule on the closing tag. The array/collection fields are
// always present here (empty until filled) and pruned to `undefined` when empty at finalisation.
interface RuleDraft {
  type: string;
  priority: number | undefined;
  stopIfTrue: boolean;
  operator: string | undefined;
  text: string | undefined;
  timePeriod: string | undefined;
  rank: number | undefined;
  stdDev: number | undefined;
  percent: boolean;
  bottom: boolean;
  aboveAverage: boolean | undefined;
  equalAverage: boolean;
  dxfId: string | undefined;
  iconSet: string | undefined;
  formulae: (string | number)[];
  cfvo: CfValueObject[];
  colors: Color[];
  color: Color | undefined;
}

/**
 * Parse the classic `<conditionalFormatting>` blocks of a worksheet into the model. Only the
 * unprefixed elements are read; the x14 extension form (`<x14:conditionalFormatting>` inside
 * `<extLst>`) is namespace-prefixed and passes straight through untouched, so its presence never
 * confuses this parser nor is it half-read into a broken rule.
 */
export function parseConditionalFormattings(xml: string): ConditionalFormatting[] {
  const blocks: ConditionalFormatting[] = [];
  let block: ConditionalFormatting | undefined;
  let draft: RuleDraft | undefined;
  let scale: ScaleKind | undefined;
  let capturingFormula = false;
  let formulaText = '';

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      if (name.includes(':')) return; // x14/xm extension elements are not the classic form
      const ln = localName(name);
      if (ln === 'conditionalFormatting') {
        block = {ref: attrs.sqref ?? '', rules: []};
      } else if (ln === 'cfRule' && block !== undefined) {
        // A rule with no operands (e.g. duplicateValues) is a self-closing element that fires no
        // close event, so it must be finalised here; one with children waits for its </cfRule>.
        if (selfClosing) {
          block.rules.push(finalizeRule(newDraft(attrs)));
        } else {
          draft = newDraft(attrs);
          scale = undefined;
        }
      } else if (draft !== undefined && (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet')) {
        scale = ln;
        if (ln === 'iconSet' && attrs.iconSet !== undefined) draft.iconSet = attrs.iconSet;
      } else if (draft !== undefined && ln === 'cfvo') {
        draft.cfvo.push(parseCfvo(attrs));
      } else if (draft !== undefined && ln === 'color') {
        const color = parseColor(attrs);
        if (scale === 'dataBar') draft.color = color;
        else draft.colors.push(color);
      } else if (draft !== undefined && ln === 'formula') {
        capturingFormula = true;
        formulaText = '';
      }
    },
    onText(chunk) {
      if (capturingFormula) formulaText += chunk;
    },
    onClose(name) {
      if (name.includes(':')) return;
      const ln = localName(name);
      if (ln === 'formula' && capturingFormula) {
        if (draft !== undefined) draft.formulae.push(coerceFormula(formulaText));
        capturingFormula = false;
      } else if (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet') {
        scale = undefined;
      } else if (ln === 'cfRule' && draft !== undefined) {
        if (block !== undefined) block.rules.push(finalizeRule(draft));
        draft = undefined;
      } else if (ln === 'conditionalFormatting' && block !== undefined) {
        blocks.push(block);
        block = undefined;
      }
    },
  });
  return blocks;
}

/**
 * Extract the differential-style (`<dxf>`) fragments from styles.xml, each verbatim. Preserving the
 * raw XML — rather than reparsing and re-serialising — is what keeps a foreign dxf's number format a
 * real format code on re-write instead of a coerced `"[object Object]"`, and keeps every conditional
 * formatting's `dxfId` index pointing at the same style it did in the source file.
 */
export function parseDxfs(stylesXml: string): string[] {
  const block = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(stylesXml);
  if (block === null) return [];
  const inner = block[1] ?? '';
  return [...inner.matchAll(/<dxf\b[^>]*>[\s\S]*?<\/dxf>|<dxf\b[^>]*\/>/g)].map(m => m[0] ?? '');
}

function newDraft(attrs: Record<string, string>): RuleDraft {
  return {
    type: attrs.type ?? '',
    priority: attrs.priority !== undefined ? Number(attrs.priority) : undefined,
    stopIfTrue: attrs.stopIfTrue === '1' || attrs.stopIfTrue === 'true',
    operator: attrs.operator,
    text: attrs.text,
    timePeriod: attrs.timePeriod,
    rank: attrs.rank !== undefined ? Number(attrs.rank) : undefined,
    stdDev: attrs.stdDev !== undefined ? Number(attrs.stdDev) : undefined,
    percent: attrs.percent === '1' || attrs.percent === 'true',
    bottom: attrs.bottom === '1' || attrs.bottom === 'true',
    // aboveAverage defaults to true in OOXML; only an explicit "0" means below-average.
    aboveAverage: attrs.aboveAverage === undefined ? undefined : attrs.aboveAverage !== '0',
    equalAverage: attrs.equalAverage === '1' || attrs.equalAverage === 'true',
    dxfId: attrs.dxfId,
    iconSet: undefined,
    formulae: [],
    cfvo: [],
    colors: [],
    color: undefined,
  };
}

function finalizeRule(draft: RuleDraft): ConditionalFormattingRule {
  const rule: ConditionalFormattingRule = {type: draft.type};
  if (draft.priority !== undefined) rule.priority = draft.priority;
  if (draft.stopIfTrue) rule.stopIfTrue = true;
  if (draft.operator !== undefined) rule.operator = draft.operator;
  if (draft.text !== undefined) rule.text = draft.text;
  if (draft.timePeriod !== undefined) rule.timePeriod = draft.timePeriod;
  if (draft.rank !== undefined) rule.rank = draft.rank;
  if (draft.stdDev !== undefined) rule.stdDev = draft.stdDev;
  if (draft.percent) rule.percent = true;
  if (draft.bottom) rule.bottom = true;
  if (draft.aboveAverage !== undefined) rule.aboveAverage = draft.aboveAverage;
  if (draft.equalAverage) rule.equalAverage = true;
  if (draft.dxfId !== undefined) rule.dxfId = draft.dxfId;
  if (draft.iconSet !== undefined) rule.iconSet = draft.iconSet;
  if (draft.formulae.length > 0) rule.formulae = draft.formulae;
  if (draft.cfvo.length > 0) rule.cfvo = draft.cfvo;
  if (draft.colors.length > 0) rule.colors = draft.colors;
  if (draft.color !== undefined) rule.color = draft.color;
  return rule;
}

function parseCfvo(attrs: Record<string, string>): CfValueObject {
  const type = (attrs.type ?? 'num') as CfValueObject['type'];
  const cfvo: CfValueObject = {type};
  if (attrs.val !== undefined) {
    // A `formula` anchor's value is an expression and stays a string; the rest are numeric.
    cfvo.value = type === 'formula' ? attrs.val : coerceNumber(attrs.val);
  }
  return cfvo;
}

function parseColor(attrs: Record<string, string>): Color {
  return {
    ...(attrs.rgb !== undefined ? {argb: attrs.rgb} : {}),
    ...(attrs.theme !== undefined ? {theme: Number(attrs.theme)} : {}),
    ...(attrs.tint !== undefined ? {tint: Number(attrs.tint)} : {}),
    ...(attrs.indexed !== undefined ? {indexed: Number(attrs.indexed)} : {}),
  };
}

// A cellIs bound / expression predicate: a bare numeric literal reads back as a number so it
// round-trips as one, while a cell reference or expression keeps its verbatim text.
function coerceFormula(text: string): string | number {
  const trimmed = text.trim();
  return /^-?\d+(?:\.\d+)?$/.test(trimmed) ? Number(trimmed) : text;
}

function coerceNumber(value: string): number | string {
  const n = Number(value);
  return Number.isFinite(n) && value.trim() !== '' ? n : value;
}
