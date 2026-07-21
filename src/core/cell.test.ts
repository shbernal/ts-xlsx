import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Cell} from './cell.ts';

test('style getter carries only the facets actually set, mirroring the per-facet getters', () => {
  const cell = new Cell(1, 1);
  assert.deepEqual(cell.style, {}, 'an untouched cell reports no facets');
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  cell.numFmt = '0.00%';
  assert.deepEqual(cell.style, {fill: cell.fill, numFmt: '0.00%'});
});

test('style setter lays each facet it carries onto the cell, readable back through the per-facet getters', () => {
  const cell = new Cell(1, 1);
  cell.style = {numFmt: '0.00', border: {top: {style: 'thin'}}};
  assert.equal(cell.numFmt, '0.00');
  assert.deepEqual(cell.border, {top: {style: 'thin'}});
});

test('style setter composes with prior per-facet sets rather than clearing them wholesale', () => {
  const cell = new Cell(1, 1);
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  cell.style = {numFmt: '0.00'};
  assert.equal(cell.numFmt, '0.00', 'the new facet lands');
  assert.deepEqual(
    cell.fill,
    {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}},
    'the facet style omits is untouched, not cleared',
  );
});
