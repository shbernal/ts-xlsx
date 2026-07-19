// Serialising conditional formatting to the classic `<conditionalFormatting>` worksheet element.
//
// Each block names its target range(s) in a `sqref` attribute and holds one or more `<cfRule>`
// children. A rule's shape depends on its type: a `dataBar`/`colorScale`/`iconSet` carries a scale
// element (its `<cfvo>` anchors and colours), while a `cellIs`/`expression`/`top10`/… carries its
// operands as `<formula>` children and points at a differential style by `dxfId`. A rule the library
// does not model in depth still round-trips its attributes, so nothing is silently dropped on save.
//
// A data bar's richer facets — its gradient fill, its negative-value fill colour, its axis colour —
// have no home in the classic `<dataBar>` element; Excel stores them only in the 2009 x14 extension.
// So a data-bar rule carrying any of them is written twice: the classic element (its anchors and bar
// colour, understood by every consumer) plus an `<x14:dataBar>` in the worksheet `<extLst>` carrying
// the extras, the two linked by a shared id. The reader folds the extension back onto the classic
// rule, so the gradient flag and the two extra colours survive a round-trip rather than being dropped.

import type {
  CfValueObject,
  ConditionalFormatting,
  ConditionalFormattingRule,
} from '../../core/conditional-formatting.ts';
import type {Color} from '../../core/style.ts';
import {colorAttrs, type StyleRegistry} from './styles.ts';
import {escapeAttr, escapeText} from './xml.ts';
import {localName, parseXml} from './xml-read.ts';

// Excel's default data bar when the author supplies none: a min/max anchor pair and its standard blue.
const DEFAULT_DATABAR_CFVO: readonly CfValueObject[] = [{type: 'min'}, {type: 'max'}];
const DEFAULT_DATABAR_COLOR: Color = {argb: 'FF638EC6'};

// The 2009 extension namespaces and the two `<ext>` uris a data bar uses: one scopes the worksheet's
// x14 conditional formattings, the other scopes the `<x14:id>` link a classic cfRule carries to name
// its extension. Declared inline exactly as Excel writes them, so no worksheet-root xmlns is needed.
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
const DATABAR_LINK_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';

// A data bar needs the x14 extension only when it carries a facet the classic element cannot express.
// A plain data bar (anchors and bar colour alone) stays classic-only, so an unadorned rule never
// fabricates an empty extension block.
function needsDataBarExt(rule: ConditionalFormattingRule): boolean {
  return (
    rule.gradient !== undefined ||
    rule.negativeFillColor !== undefined ||
    rule.axisColor !== undefined
  );
}

// The synthetic id linking a classic cfRule to its x14 extension. Excel uses a random GUID; any unique
// token that matches on both ends works, so a deterministic per-sheet index keeps the output stable
// and testable. The classic and extension passes iterate rules identically, so the Nth extended data
// bar gets the same id on both ends.
function dataBarExtGuid(index: number): string {
  return `{00000000-0000-0000-0000-${String(index + 1).padStart(12, '0')}}`;
}

// The scale rules render a built-in visual and carry no differential style; the rest apply a dxf.
const SCALE_TYPES = new Set(['dataBar', 'colorScale', 'iconSet']);

/**
 * Serialise every conditional formatting on a sheet into its `<conditionalFormatting>` blocks, in
 * insertion order. Returns '' when the sheet has none. A rule missing a `priority` is assigned the
 * next free one, so the output always satisfies Excel's requirement that every cfRule carry one.
 */
export function conditionalFormattingsXml(
  formattings: readonly ConditionalFormatting[],
  styles: StyleRegistry,
): string {
  if (formattings.length === 0) return '';
  const priority = {next: 1};
  // A shared counter assigns each extended data bar its link id in block/rule order — the same order
  // {@link conditionalFormattingsExtXml} walks, so the two ends of a link always agree.
  const extGuid = {next: 0};
  return formattings.map((cf) => blockXml(cf, styles, priority, extGuid)).join('');
}

/**
 * The worksheet `<extLst>` `<ext>` carrying the x14 data-bar extensions, or '' when no data bar needs
 * one. Walks the rules in the same order as {@link conditionalFormattingsXml}, so the Nth extended
 * data bar's `<x14:cfRule id>` matches the `<x14:id>` its classic cfRule carries. Emitted bare (no
 * `<extLst>` wrapper) so the worksheet serialiser can gather it into a single `<extLst>` beside the
 * data-validation extension.
 */
export function conditionalFormattingsExtXml(
  formattings: readonly ConditionalFormatting[],
): string {
  const items: string[] = [];
  let index = 0;
  for (const cf of formattings) {
    for (const rule of cf.rules) {
      if (rule.type === 'dataBar' && needsDataBarExt(rule)) {
        items.push(x14DataBarXml(cf.ref, rule, dataBarExtGuid(index)));
        index += 1;
      }
    }
  }
  if (items.length === 0) return '';
  return (
    `<ext uri="${CF_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:conditionalFormattings>${items.join('')}</x14:conditionalFormattings></ext>`
  );
}

// One `<x14:conditionalFormatting>`: an `<x14:cfRule type="dataBar" id>` mirroring the classic anchors
// as `<x14:cfvo>` and adding the facets the classic element cannot carry (gradient, negative-fill and
// axis colours), with the target range in an `<xm:sqref>` child — the shape Excel writes.
function x14DataBarXml(ref: string, rule: ConditionalFormattingRule, guid: string): string {
  const cfvo = rule.cfvo && rule.cfvo.length > 0 ? rule.cfvo : DEFAULT_DATABAR_CFVO;
  const anchors = cfvo.map(x14CfvoXml).join('');
  const gradient = rule.gradient !== undefined ? ` gradient="${rule.gradient ? 1 : 0}"` : '';
  const negative =
    rule.negativeFillColor !== undefined
      ? `<x14:negativeFillColor ${colorAttrs(rule.negativeFillColor)}/>`
      : '';
  const axis = rule.axisColor !== undefined ? `<x14:axisColor ${colorAttrs(rule.axisColor)}/>` : '';
  return (
    `<x14:conditionalFormatting xmlns:xm="${XM_NS}">` +
    `<x14:cfRule type="dataBar" id="${guid}">` +
    `<x14:dataBar${gradient}>${anchors}${negative}${axis}</x14:dataBar>` +
    `</x14:cfRule><xm:sqref>${escapeText(ref)}</xm:sqref></x14:conditionalFormatting>`
  );
}

// An x14 scale anchor. A `min`/`max` carries no value and self-closes; the rest wrap their value in an
// `<xm:f>` (the extension form stores every anchor value as a formula).
function x14CfvoXml(cfvo: CfValueObject): string {
  const type = escapeAttr(cfvo.type);
  if (cfvo.value === undefined) return `<x14:cfvo type="${type}"/>`;
  return `<x14:cfvo type="${type}"><xm:f>${escapeText(String(cfvo.value))}</xm:f></x14:cfvo>`;
}

// The `<extLst>` a classic data-bar cfRule carries to name its x14 extension by shared id.
function cfRuleExtLinkXml(guid: string): string {
  return (
    `<extLst><ext uri="${DATABAR_LINK_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:id>${guid}</x14:id></ext></extLst>`
  );
}

function blockXml(
  cf: ConditionalFormatting,
  styles: StyleRegistry,
  priority: {next: number},
  extGuid: {next: number},
): string {
  const rules = cf.rules.map((rule) => ruleXml(rule, styles, priority, extGuid)).join('');
  return `<conditionalFormatting sqref="${escapeAttr(cf.ref)}">${rules}</conditionalFormatting>`;
}

function ruleXml(
  rule: ConditionalFormattingRule,
  styles: StyleRegistry,
  priority: {next: number},
  extGuid: {next: number},
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

  let body = SCALE_TYPES.has(rule.type) ? scaleXml(rule) : formulaeXml(rule.formulae);
  // A data bar with x14-only facets links to its extension by a freshly allocated id; the extension
  // itself rides in the worksheet <extLst>. The link is the cfRule's last child, after the dataBar.
  if (rule.type === 'dataBar' && needsDataBarExt(rule)) {
    body += cfRuleExtLinkXml(dataBarExtGuid(extGuid.next++));
  }
  return body === ''
    ? `<cfRule ${attrs.join(' ')}/>`
    : `<cfRule ${attrs.join(' ')}>${body}</cfRule>`;
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
  return formulae.map((f) => `<formula>${escapeText(formulaText(f))}</formula>`).join('');
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
// empty element. The gradient flag and the negative-fill/axis colours have no home in this classic
// element; they ride in the x14 extension (see {@link conditionalFormattingsExtXml}), linked from the
// cfRule that wraps this by a shared id.
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
  const colors = (rule.colors ?? []).map((c) => `<color ${colorAttrs(c)}/>`).join('');
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
  // The `<x14:id>` a data-bar cfRule carries to name its extension. Transient — it links this rule to
  // its `<x14:dataBar>` during parsing and is dropped once the extension's facets are folded in.
  x14Id: string | undefined;
}

// The facets an `<x14:dataBar>` adds over the classic element, gathered by the id its `<x14:cfRule>`
// carries so they can be matched to the classic rule that links to the same id.
interface DataBarExt {
  gradient: boolean | undefined;
  negativeFillColor: Color | undefined;
  axisColor: Color | undefined;
}

/**
 * Parse a worksheet's conditional formatting into the model. The classic `<conditionalFormatting>`
 * blocks supply every rule; the x14 extension (`<x14:conditionalFormatting>` inside `<extLst>`) is
 * read only to enrich a classic data bar with the facets the classic element cannot carry — the
 * gradient flag and the negative-fill and axis colours — matched by the shared id the two ends link
 * on. An extension rule with no classic counterpart (a rule that lives only in x14) is ignored, so it
 * is never half-read into a broken classic rule.
 */
export function parseConditionalFormattings(xml: string): ConditionalFormatting[] {
  const blocks: ConditionalFormatting[] = [];
  let block: ConditionalFormatting | undefined;
  let draft: RuleDraft | undefined;
  let scale: ScaleKind | undefined;
  let capturingFormula = false;
  let formulaText = '';

  // Classic data-bar rules that named an extension, paired with the id they linked on, plus the
  // extensions gathered from the worksheet <extLst>. The two are married after the pass — the
  // extension always follows the classic blocks in document order, so it is known by then.
  const linked: {rule: ConditionalFormattingRule; id: string}[] = [];
  const extById = new Map<string, DataBarExt>();
  let x14Ext: DataBarExt | undefined;
  let x14ExtId: string | undefined;
  let capturingX14Id = false;
  let x14IdText = '';

  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      const ln = localName(name);
      if (name.includes(':')) {
        // The `<x14:id>` a classic data bar carries to name its extension: capture its text into the
        // open draft. The rest are the worksheet extension's own elements.
        if (ln === 'id' && draft !== undefined) {
          capturingX14Id = true;
          x14IdText = '';
        } else if (ln === 'cfRule') {
          x14Ext = attrs.type === 'dataBar' && attrs.id !== undefined ? emptyExt() : undefined;
          x14ExtId = attrs.id;
        } else if (x14Ext !== undefined && ln === 'dataBar') {
          // gradient defaults to true in the x14 schema, so an absent attribute reads as a gradient.
          x14Ext.gradient = attrs.gradient !== '0';
        } else if (x14Ext !== undefined && ln === 'negativeFillColor') {
          x14Ext.negativeFillColor = parseColor(attrs);
        } else if (x14Ext !== undefined && ln === 'axisColor') {
          x14Ext.axisColor = parseColor(attrs);
        }
        return;
      }
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
      } else if (
        draft !== undefined &&
        (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet')
      ) {
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
      if (capturingX14Id) x14IdText += chunk;
    },
    onClose(name) {
      const ln = localName(name);
      if (name.includes(':')) {
        if (ln === 'id' && capturingX14Id) {
          if (draft !== undefined) draft.x14Id = x14IdText;
          capturingX14Id = false;
        } else if (ln === 'cfRule' && x14Ext !== undefined && x14ExtId !== undefined) {
          extById.set(x14ExtId, x14Ext);
          x14Ext = undefined;
          x14ExtId = undefined;
        }
        return;
      }
      if (ln === 'formula' && capturingFormula) {
        if (draft !== undefined) draft.formulae.push(coerceFormula(formulaText));
        capturingFormula = false;
      } else if (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet') {
        scale = undefined;
      } else if (ln === 'cfRule' && draft !== undefined) {
        const rule = finalizeRule(draft);
        if (block !== undefined) block.rules.push(rule);
        if (draft.x14Id !== undefined) linked.push({rule, id: draft.x14Id});
        draft = undefined;
      } else if (ln === 'conditionalFormatting' && block !== undefined) {
        blocks.push(block);
        block = undefined;
      }
    },
  });

  for (const {rule, id} of linked) {
    const ext = extById.get(id);
    if (ext === undefined) continue;
    if (ext.gradient !== undefined) rule.gradient = ext.gradient;
    if (ext.negativeFillColor !== undefined) rule.negativeFillColor = ext.negativeFillColor;
    if (ext.axisColor !== undefined) rule.axisColor = ext.axisColor;
  }
  return blocks;
}

function emptyExt(): DataBarExt {
  return {gradient: undefined, negativeFillColor: undefined, axisColor: undefined};
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
  return [...inner.matchAll(/<dxf\b[^>]*>[\s\S]*?<\/dxf>|<dxf\b[^>]*\/>/g)].map((m) => m[0] ?? '');
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
    x14Id: undefined,
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
