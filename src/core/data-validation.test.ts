import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from './workbook.ts';

test('a range validation is reported on every cell the range covers, not just the first', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('A1:A3', {type: 'whole', operator: 'between', formulae: [0, 9]});

  for (const ref of ['A1', 'A2', 'A3']) {
    const dv = sheet.dataValidationAt(ref);
    assert.ok(dv, `${ref} should carry the range's validation`);
    assert.equal(dv.type, 'whole');
    assert.deepEqual(dv.formulae, [0, 9]);
  }
  assert.equal(sheet.dataValidationAt('A4'), undefined, 'a cell outside the range carries none');
  assert.equal(sheet.dataValidationAt('B1'), undefined, 'a cell in another column carries none');
});

test('the first added validation whose range contains a cell wins', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('A1:A3', {type: 'whole', formulae: [1]});
  sheet.addDataValidation('A2', {type: 'whole', formulae: [2]});

  assert.deepEqual(sheet.dataValidationAt('A2')?.formulae, [1], 'the earlier, broader rule wins');
});

test('a whole-column validation covers a cell far down the column without per-cell storage', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('B2:B1048576', {type: 'list', formulae: ['"a,b,c"']});

  assert.ok(sheet.dataValidationAt('B2'), 'the top of the column is covered');
  assert.ok(sheet.dataValidationAt('B1000'), 'a cell deep in the column is covered');
  assert.equal(sheet.dataValidationAt('B1'), undefined, 'a cell above the range is not covered');
  assert.equal(sheet.dataValidations.length, 1, 'the column is a single entry, not a million');
});

test('a stored validation does not alias the caller-supplied rule or its formulae', () => {
  const sheet = new Workbook().addWorksheet('S');
  const rule = {type: 'whole' as const, formulae: [1, 2]};
  sheet.addDataValidation('A1', rule);
  rule.formulae[0] = 99;

  assert.deepEqual(sheet.dataValidationAt('A1')?.formulae, [1, 2], 'the stored rule is a copy');
});

test('data validations survive a model round-trip', () => {
  const source = new Workbook().addWorksheet('src');
  source.addDataValidation('A1:A3', {type: 'whole', operator: 'between', formulae: [0, 9]});
  source.addDataValidation('B2', {type: 'list', formulae: ['"x,y"']});

  const dest = new Workbook().addWorksheet('dst');
  dest.model = source.model;

  assert.equal(dest.dataValidations.length, 2);
  assert.deepEqual(dest.dataValidationAt('A2')?.formulae, [0, 9]);
  assert.equal(dest.dataValidationAt('B2')?.type, 'list');
});

test('the extended flag rides on the entry and survives a model round-trip', () => {
  const source = new Workbook().addWorksheet('src');
  source.addDataValidation('A1', {type: 'list', formulae: ['Sheet2!$A:$A']}, {extended: true});
  source.addDataValidation('B1', {type: 'list', formulae: ['"x,y"']});

  const [ext, std] = source.dataValidations;
  assert.equal(ext?.extended, true, 'the extended entry is tagged');
  assert.equal(std?.extended, undefined, 'a standard entry carries no flag');

  const dest = new Workbook().addWorksheet('dst');
  dest.model = source.model;
  assert.equal(dest.dataValidations[0]?.extended, true, 'the flag survives the round-trip');
  assert.equal(dest.dataValidations[1]?.extended, undefined);
});
