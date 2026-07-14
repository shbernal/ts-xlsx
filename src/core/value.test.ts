import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  type CellValue,
  ValueType,
  coerceCellValue,
  detectValueType,
  isErrorCode,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
} from './value.ts';

test('detectValueType classifies each primitive', () => {
  assert.equal(detectValueType(null), ValueType.Null);
  assert.equal(detectValueType(42), ValueType.Number);
  assert.equal(detectValueType('hi'), ValueType.String);
  assert.equal(detectValueType(true), ValueType.Boolean);
  assert.equal(detectValueType(new Date(0)), ValueType.Date);
});

test('an invalid Date is still a Date-typed value', () => {
  assert.equal(detectValueType(new Date(NaN)), ValueType.Date);
});

test('NaN and Infinity remain Number-typed — the model does not judge finiteness', () => {
  assert.equal(detectValueType(Number.NaN), ValueType.Number);
  assert.equal(detectValueType(Number.POSITIVE_INFINITY), ValueType.Number);
});

test('detectValueType classifies structural values', () => {
  assert.equal(detectValueType({error: '#REF!'}), ValueType.Error);
  assert.equal(detectValueType({formula: 'A1+B1'}), ValueType.Formula);
  assert.equal(detectValueType({formula: 'A1', result: 3}), ValueType.Formula);
  assert.equal(detectValueType({sharedFormula: 'A1', result: 3}), ValueType.Formula);
  assert.equal(detectValueType({richText: [{text: 'a'}]}), ValueType.RichText);
  assert.equal(detectValueType({hyperlink: 'https://x', text: 'x'}), ValueType.Hyperlink);
});

test('a hyperlink whose text is rich still classifies as Hyperlink, not RichText', () => {
  const value: CellValue = {hyperlink: 'https://x', text: {richText: [{text: 'x'}]}};
  assert.equal(detectValueType(value), ValueType.Hyperlink);
});

test('detectValueType throws on an unrecognised object shape', () => {
  assert.throws(() => detectValueType({nonsense: true} as unknown as CellValue), TypeError);
});

test('type guards discriminate the structural shapes', () => {
  assert.ok(isFormulaValue({formula: 'A1'}));
  assert.ok(!isFormulaValue({sharedFormula: 'A1'}));
  assert.ok(isSharedFormulaValue({sharedFormula: 'A1'}));
  assert.ok(isRichTextValue({richText: []}));
  assert.ok(isHyperlinkValue({hyperlink: 'u', text: 't'}));
});

test('isErrorCode recognises the canonical literals only', () => {
  assert.ok(isErrorCode('#DIV/0!'));
  assert.ok(isErrorCode('#N/A'));
  assert.ok(!isErrorCode('#WHATEVER!'));
  assert.ok(!isErrorCode('42'));
});

test('coerceCellValue maps undefined to the empty cell', () => {
  assert.equal(coerceCellValue(undefined), null);
});

test('coerceCellValue preserves a numeric-looking string as a string', () => {
  const coerced = coerceCellValue('1000.80');
  assert.equal(coerced, '1000.80');
  assert.equal(detectValueType(coerced), ValueType.String);
});

test('coerceCellValue passes valid values through unchanged', () => {
  const date = new Date(0);
  assert.equal(coerceCellValue(date), date);
  const formula = {formula: 'A1'};
  assert.equal(coerceCellValue(formula), formula);
});

test('coerceCellValue rejects an unrecognised shape at the assignment site', () => {
  assert.throws(() => coerceCellValue({bogus: 1} as unknown as CellValue), TypeError);
});
