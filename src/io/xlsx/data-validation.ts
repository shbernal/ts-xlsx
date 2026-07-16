// Data validation I/O — the sheet-level `<dataValidations>` element and the reader that folds each
// rule back onto its range.
//
// A validation stores its operands in `<formula1>`/`<formula2>` child elements and its target in a
// `sqref` attribute. The operand text carries NO leading '=' (the '=' is an authoring convention, not
// part of the stored formula), so the writer strips one and the reader keeps whatever it finds. A
// numeric-typed rule's literal operand parses to a number; a cell reference, defined name, or list
// source keeps its string, so a reference is never coerced to NaN and lost.
//
// The extended `<x14:dataValidation>` form (2009 extension schema, used for cross-sheet and some
// whole-column list validations) is a separate, larger concern: this reader only handles the standard
// unprefixed element and safely ignores the x14 block rather than mis-parsing it.

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

/** The `<dataValidations>` element, or '' when the sheet declares none — so a validation-free sheet
 * stays byte-clean. */
export function dataValidationsXml(entries: readonly DataValidationEntry[]): string {
  if (entries.length === 0) return '';
  const items = entries.map(({sqref, rule}) => dataValidationXml(sqref, rule)).join('');
  return `<dataValidations count="${entries.length}">${items}</dataValidations>`;
}

// Attribute order follows CT_DataValidation: type, errorStyle, operator, allowBlank,
// showInputMessage, showErrorMessage, errorTitle, error, promptTitle, prompt, sqref (last).
function dataValidationXml(sqref: string, rule: DataValidation): string {
  const attrs =
    ` type="${rule.type}"` +
    (rule.errorStyle !== undefined ? ` errorStyle="${rule.errorStyle}"` : '') +
    (rule.operator !== undefined ? ` operator="${rule.operator}"` : '') +
    (rule.allowBlank ? ' allowBlank="1"' : '') +
    (rule.showInputMessage ? ' showInputMessage="1"' : '') +
    (rule.showErrorMessage ? ' showErrorMessage="1"' : '') +
    (rule.errorTitle !== undefined ? ` errorTitle="${escapeAttr(rule.errorTitle)}"` : '') +
    (rule.error !== undefined ? ` error="${escapeAttr(rule.error)}"` : '') +
    (rule.promptTitle !== undefined ? ` promptTitle="${escapeAttr(rule.promptTitle)}"` : '') +
    (rule.prompt !== undefined ? ` prompt="${escapeAttr(rule.prompt)}"` : '') +
    ` sqref="${escapeAttr(sqref)}"`;

  const formulae = rule.formulae ?? [];
  const [f1, f2] = formulae;
  const body =
    (f1 !== undefined ? `<formula1>${formulaText(f1)}</formula1>` : '') +
    (f2 !== undefined ? `<formula2>${formulaText(f2)}</formula2>` : '');
  return `<dataValidation${attrs}>${body}</dataValidation>`;
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
  const {sqref, type} = attrs;
  if (sqref === undefined || type === undefined) return undefined;

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

  return {sqref, rule};
}

// A list/custom operand is always a string. A typed rule's operand is a number when it is a plain
// numeric literal, and otherwise its verbatim string — so a cell reference (`L26`) or defined name
// survives instead of being coerced to NaN.
function parseFormula(type: string, text: string): string | number {
  if (type === 'list' || type === 'custom') return text;
  const n = Number(text);
  return text.trim() !== '' && Number.isFinite(n) ? n : text;
}

/** Fold parsed validations onto a sheet, each bound to its original range. */
export function applyDataValidations(
  sheet: Worksheet,
  entries: readonly DataValidationEntry[]
): void {
  for (const {sqref, rule} of entries) sheet.addDataValidation(sqref, rule);
}
