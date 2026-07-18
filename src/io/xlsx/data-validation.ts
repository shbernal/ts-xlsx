// Data validation I/O — the sheet-level `<dataValidations>` element and the reader that folds each
// rule back onto its range.
//
// A validation stores its operands in `<formula1>`/`<formula2>` child elements and its target in a
// `sqref` attribute. The operand text carries NO leading '=' (the '=' is an authoring convention, not
// part of the stored formula), so the writer strips one and the reader keeps whatever it finds. A
// numeric-typed rule's literal operand parses to a number; a cell reference, defined name, or list
// source keeps its string, so a reference is never coerced to NaN and lost.
//
// The extended `<x14:dataValidation>` form (2009 extension schema) carries the validations a legacy
// element cannot express — chiefly a list whose source lives on another sheet. It lives in the
// worksheet `<extLst>`, keeps its target in a `<xm:sqref>` child rather than a `sqref` attribute, and
// wraps each operand in an `<xm:f>` under `<x14:formula1>`/`<x14:formula2>`. A rule read from that
// form is tagged `extended` so it is written back there; the two forms are parsed and serialised by
// prefix so neither reader mistakes one for the other.

import {
  type DataValidation,
  type DataValidationEntry,
  type DataValidationErrorStyle,
  type DataValidationOperator,
  type DataValidationType,
} from '../../core/data-validation.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {localName, parseXml} from './xml-read.ts';
import {escapeAttr, escapeText} from './xml.ts';

// The typed validations whose literal operands are numbers; `list`/`custom` operands stay strings.
const TYPED = new Set<string>(['whole', 'decimal', 'date', 'time', 'textLength']);

// The 2009 extension namespaces and the `<ext>` uri that scopes a worksheet's extended validations —
// declared inline on the elements that need them, exactly as Excel writes them, so the block is
// self-contained and the worksheet root needs no extra namespace declaration.
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const DATA_VALIDATION_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';

/** The standard `<dataValidations>` element for the rules stored in the legacy form, or '' when the
 * sheet has none of them — so a sheet with only extended (or no) validations stays byte-clean here.
 * The extended rules are emitted separately by {@link dataValidationsExtXml}. */
export function dataValidationsXml(entries: readonly DataValidationEntry[]): string {
  const standard = entries.filter(entry => !entry.extended);
  if (standard.length === 0) return '';
  const items = standard.map(({sqref, rule}) => dataValidationXml(sqref, rule)).join('');
  return `<dataValidations count="${standard.length}">${items}</dataValidations>`;
}

/** The `<ext>` carrying the extended (`<x14:dataValidation>`) rules, or '' when the sheet declares
 * none. Emitted bare (no `<extLst>` wrapper) so the worksheet serialiser can gather it into a single
 * `<extLst>` beside the conditional-formatting extension — a worksheet may carry at most one. */
export function dataValidationsExtXml(entries: readonly DataValidationEntry[]): string {
  const extended = entries.filter(entry => entry.extended);
  if (extended.length === 0) return '';
  const items = extended.map(({sqref, rule}) => extendedDataValidationXml(sqref, rule)).join('');
  return (
    `<ext uri="${DATA_VALIDATION_EXT_URI}" xmlns:x14="${X14_NS}">` +
    `<x14:dataValidations count="${extended.length}" xmlns:xm="${XM_NS}">${items}</x14:dataValidations>` +
    '</ext>'
  );
}

// The shared attributes of a validation, in CT_DataValidation order: type, errorStyle, operator,
// allowBlank, showInputMessage, showErrorMessage, errorTitle, error, promptTitle, prompt. The target
// range differs between the two forms (a `sqref` attribute vs an `<xm:sqref>` child), so it is not
// part of this shared prefix.
function ruleAttrs(rule: DataValidation): string {
  return (
    ` type="${rule.type}"` +
    (rule.errorStyle !== undefined ? ` errorStyle="${rule.errorStyle}"` : '') +
    (rule.operator !== undefined ? ` operator="${rule.operator}"` : '') +
    (rule.allowBlank ? ' allowBlank="1"' : '') +
    (rule.showInputMessage ? ' showInputMessage="1"' : '') +
    (rule.showErrorMessage ? ' showErrorMessage="1"' : '') +
    (rule.errorTitle !== undefined ? ` errorTitle="${escapeAttr(rule.errorTitle)}"` : '') +
    (rule.error !== undefined ? ` error="${escapeAttr(rule.error)}"` : '') +
    (rule.promptTitle !== undefined ? ` promptTitle="${escapeAttr(rule.promptTitle)}"` : '') +
    (rule.prompt !== undefined ? ` prompt="${escapeAttr(rule.prompt)}"` : '')
  );
}

// The standard element: shared attributes, then `sqref` last, then `<formula1>`/`<formula2>` bodies.
function dataValidationXml(sqref: string, rule: DataValidation): string {
  const [f1, f2] = operands(rule);
  const body =
    (f1 !== undefined ? `<formula1>${formulaText(f1)}</formula1>` : '') +
    (f2 !== undefined ? `<formula2>${formulaText(f2)}</formula2>` : '');
  return `<dataValidation${ruleAttrs(rule)} sqref="${escapeAttr(sqref)}">${body}</dataValidation>`;
}

// A rule's two operands with any non-finite numeric bound dropped: a NaN/±Infinity operand (e.g. a
// date validation whose bound failed to coerce to a serial) has no OOXML representation, so it is
// omitted rather than serialised as the literal "NaN" — the same graceful degradation a non-finite
// cell value gets.
function operands(rule: DataValidation): [string | number | undefined, string | number | undefined] {
  const drop = (v: string | number | undefined): string | number | undefined =>
    typeof v === 'number' && !Number.isFinite(v) ? undefined : v;
  const [f1, f2] = rule.formulae ?? [];
  return [drop(f1), drop(f2)];
}

// The extended element: same shared attributes, but each operand wraps in `<x14:formula1><xm:f>…` and
// the target range is an `<xm:sqref>` child that follows the formulae. The `xr:uid` Excel adds is
// revision metadata it regenerates freely, so it is not modelled or re-emitted.
function extendedDataValidationXml(sqref: string, rule: DataValidation): string {
  const [f1, f2] = operands(rule);
  const body =
    (f1 !== undefined ? `<x14:formula1><xm:f>${formulaText(f1)}</xm:f></x14:formula1>` : '') +
    (f2 !== undefined ? `<x14:formula2><xm:f>${formulaText(f2)}</xm:f></x14:formula2>` : '') +
    `<xm:sqref>${escapeText(sqref)}</xm:sqref>`;
  return `<x14:dataValidation${ruleAttrs(rule)}>${body}</x14:dataValidation>`;
}

// A number serialises as its literal; a string is stripped of exactly one leading '=' (the authoring
// convention OOXML omits) and escaped as element text.
function formulaText(value: string | number): string {
  if (typeof value === 'number') return String(value);
  const stripped = value.startsWith('=') ? value.slice(1) : value;
  return escapeText(stripped);
}

/** Parse every standard `<dataValidation>` out of a worksheet part into range-bound rules. */
export function parseDataValidations(xml: string): DataValidationEntry[] {
  const entries: DataValidationEntry[] = [];
  let current: {attrs: Record<string, string>; formulae: string[]} | undefined;
  let slot: number | undefined;

  parseXml(xml, {
    onOpen(name, attrs) {
      const ln = localName(name);
      // Only the standard, unprefixed element — an `x14:dataValidation` is left for the extended path.
      if (ln === 'dataValidation' && !name.includes(':')) {
        current = {attrs, formulae: []};
      } else if (current !== undefined && ln === 'formula1') {
        slot = 0;
        current.formulae[0] = '';
      } else if (current !== undefined && ln === 'formula2') {
        slot = 1;
        current.formulae[1] = '';
      }
    },
    onText(text) {
      if (current !== undefined && slot !== undefined) {
        current.formulae[slot] = (current.formulae[slot] ?? '') + text;
      }
    },
    onClose(name) {
      const ln = localName(name);
      if (ln === 'formula1' || ln === 'formula2') {
        slot = undefined;
        return;
      }
      if (ln === 'dataValidation' && current !== undefined) {
        const built = buildEntry(current.attrs, current.formulae);
        if (built !== undefined) entries.push(built);
        current = undefined;
      }
    },
  });
  return entries;
}

function buildEntry(
  attrs: Record<string, string>,
  formulae: readonly string[]
): DataValidationEntry | undefined {
  const {sqref} = attrs;
  if (sqref === undefined) return undefined;
  const rule = buildRule(attrs, formulae);
  return rule === undefined ? undefined : {sqref, rule};
}

// The rule carried by a validation element of either form: its attributes decide the type, operator,
// flags, and messages; its `<formula1>`/`<formula2>` operands become `formulae`. The target range is
// supplied separately by each form's caller, so it is not read here.
function buildRule(
  attrs: Record<string, string>,
  formulae: readonly string[]
): DataValidation | undefined {
  const {type} = attrs;
  if (type === undefined) return undefined;

  const rule: DataValidation = {type: type as DataValidationType};
  if (attrs.operator !== undefined) {
    rule.operator = attrs.operator as DataValidationOperator;
  } else if (TYPED.has(type)) {
    // Excel omits `operator="between"` because it is the default for a typed rule; restore it so a
    // reader sees the operator the rule actually enforces.
    rule.operator = 'between';
  }
  if (attrs.allowBlank === '1') rule.allowBlank = true;
  if (attrs.showInputMessage === '1') rule.showInputMessage = true;
  if (attrs.showErrorMessage === '1') rule.showErrorMessage = true;
  if (attrs.errorStyle !== undefined) rule.errorStyle = attrs.errorStyle as DataValidationErrorStyle;
  if (attrs.error !== undefined) rule.error = attrs.error;
  if (attrs.errorTitle !== undefined) rule.errorTitle = attrs.errorTitle;
  if (attrs.prompt !== undefined) rule.prompt = attrs.prompt;
  if (attrs.promptTitle !== undefined) rule.promptTitle = attrs.promptTitle;

  const parsed = formulae
    .filter((f): f is string => f !== undefined)
    .map((f) => parseFormula(type, f));
  if (parsed.length > 0) rule.formulae = parsed;

  return rule;
}

/** Parse every extended `<x14:dataValidation>` out of a worksheet's `<extLst>` into range-bound
 * rules tagged `extended`, so a cross-sheet or whole-column list validation Excel stored only in the
 * 2009 extension form is read back rather than dropped. The standard parser ignores these (they are
 * prefixed); this one, symmetrically, handles only the prefixed elements. */
export function parseExtendedDataValidations(xml: string): DataValidationEntry[] {
  const entries: DataValidationEntry[] = [];
  let current: {attrs: Record<string, string>; formulae: string[]; sqref: string} | undefined;
  // Which operand an `<xm:f>` feeds (set by the enclosing `<x14:formula1>`/`<x14:formula2>`), and
  // which child element's text is currently being gathered.
  let slot: number | undefined;
  let capture: 'formula' | 'sqref' | undefined;
  let text = '';

  parseXml(xml, {
    onOpen(name, attrs) {
      const ln = localName(name);
      const prefixed = name.includes(':');
      // A `<x14:dataValidation>`; its attributes (type, flags, messages) build the rule.
      if (ln === 'dataValidation' && prefixed) {
        current = {attrs, formulae: [], sqref: ''};
      } else if (current !== undefined && prefixed && ln === 'formula1') {
        slot = 0;
      } else if (current !== undefined && prefixed && ln === 'formula2') {
        slot = 1;
      } else if (current !== undefined && ln === 'f') {
        capture = 'formula';
        text = '';
      } else if (current !== undefined && ln === 'sqref') {
        capture = 'sqref';
        text = '';
      }
    },
    onText(chunk) {
      if (capture !== undefined) text += chunk;
    },
    onClose(name) {
      const ln = localName(name);
      if (ln === 'f' && capture === 'formula') {
        if (current !== undefined && slot !== undefined) current.formulae[slot] = text;
        capture = undefined;
      } else if (ln === 'sqref' && capture === 'sqref') {
        if (current !== undefined) current.sqref = text;
        capture = undefined;
      } else if ((ln === 'formula1' || ln === 'formula2') && name.includes(':')) {
        slot = undefined;
      } else if (ln === 'dataValidation' && name.includes(':') && current !== undefined) {
        const built = buildExtendedEntry(current.attrs, current.formulae, current.sqref);
        if (built !== undefined) entries.push(built);
        current = undefined;
      }
    },
  });
  return entries;
}

function buildExtendedEntry(
  attrs: Record<string, string>,
  formulae: readonly string[],
  sqref: string
): DataValidationEntry | undefined {
  if (sqref === '') return undefined;
  const rule = buildRule(attrs, formulae);
  return rule === undefined ? undefined : {sqref, rule, extended: true};
}

// A list/custom operand is always a string. A typed rule's operand is a number when it is a plain
// numeric literal, and otherwise its verbatim string — so a cell reference (`L26`) or defined name
// survives instead of being coerced to NaN.
function parseFormula(type: string, text: string): string | number {
  if (type === 'list' || type === 'custom') return text;
  const n = Number(text);
  return text.trim() !== '' && Number.isFinite(n) ? n : text;
}

/** Fold parsed validations onto a sheet, each bound to its original range and carrying its form: an
 * `extended` entry is re-attached as extended so a round-trip writes it back to the x14 block. */
export function applyDataValidations(
  sheet: Worksheet,
  entries: readonly DataValidationEntry[]
): void {
  for (const {sqref, rule, extended} of entries) {
    sheet.addDataValidation(sqref, rule, extended ? {extended: true} : {});
  }
}
