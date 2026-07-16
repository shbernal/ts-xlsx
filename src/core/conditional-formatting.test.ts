import assert from 'node:assert/strict';
import {test} from 'node:test';

import {cloneConditionalFormatting} from './conditional-formatting.ts';
import {Workbook} from './workbook.ts';

test('a stored conditional formatting does not alias the caller-supplied object or its nested arrays', () => {
  const sheet = new Workbook().addWorksheet('S');
  const cf = {
    ref: 'A1:A3',
    rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [3], cfvo: [{type: 'num' as const, value: 0}]}],
  };
  sheet.addConditionalFormatting(cf);
  cf.rules[0]!.formulae[0] = 99;
  cf.rules[0]!.cfvo[0]!.value = 88;
  cf.ref = 'ZZ1';

  const stored = sheet.conditionalFormattings[0];
  assert.equal(stored?.ref, 'A1:A3', 'the ref is a copy');
  assert.deepEqual(stored?.rules[0]?.formulae, [3], 'the formulae array is a copy');
  assert.equal(stored?.rules[0]?.cfvo?.[0]?.value, 0, 'the cfvo objects are copies');
});

test('cloneConditionalFormatting deep-copies rules, formulae, cfvo, colours, and the differential style', () => {
  const original = {
    ref: 'B2:B10',
    rules: [
      {
        type: 'colorScale',
        cfvo: [{type: 'min' as const}, {type: 'max' as const}],
        colors: [{argb: 'FFFF0000'}, {argb: 'FF00FF00'}],
        style: {fill: {type: 'pattern' as const, pattern: 'solid' as const, bgColor: {argb: 'FF0000FF'}}},
      },
    ],
  };
  const copy = cloneConditionalFormatting(original);
  copy.rules[0]!.colors![0] = {argb: 'FFFFFFFF'};
  copy.rules[0]!.cfvo![0] = {type: 'num', value: 1};

  assert.deepEqual(original.rules[0]!.colors[0], {argb: 'FFFF0000'}, 'colours are not aliased');
  assert.deepEqual(original.rules[0]!.cfvo[0], {type: 'min'}, 'cfvo entries are not aliased');
  assert.notEqual(copy.rules[0]!.style, original.rules[0]!.style, 'the style is a fresh object');
});

test('conditional formattings survive a worksheet model round-trip', () => {
  const source = new Workbook().addWorksheet('src');
  source.addConditionalFormatting({
    ref: 'A1:A3',
    rules: [{type: 'dataBar', color: {argb: 'FF638EC6'}, cfvo: [{type: 'num', value: 0}, {type: 'num', value: 1}]}],
  });
  source.addConditionalFormatting({
    ref: 'C1:C9',
    rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 2}],
  });

  const dest = new Workbook().addWorksheet('dst');
  dest.model = source.model;

  assert.equal(dest.conditionalFormattings.length, 2, 'both blocks survive');
  assert.equal(dest.conditionalFormattings[0]?.rules[0]?.type, 'dataBar');
  assert.deepEqual(dest.conditionalFormattings[0]?.rules[0]?.cfvo?.map(v => v.value), [0, 1]);
  assert.equal(dest.conditionalFormattings[1]?.rules[0]?.priority, 2);
});
