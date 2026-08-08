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

// --- rich-text runs inheriting the cell font ------------------------------------------------------

test('setRichText composes each run over the cell font, so a run states only what it changes', () => {
  // A run's <rPr> is a complete character format: a facet it omits falls back to the workbook
  // default font, NOT to the cell's. Verified against Excel — a cell set to Courier New 16 whose
  // first run carries only <b/> renders that run in the workbook default face. So authoring
  // `{bold: true}` beside a styled cell silently loses the face unless the face is composed in.
  const cell = new Cell(1, 1);
  cell.font = {name: 'Courier New', size: 16, color: {theme: 1}};
  cell.setRichText([{text: 'Note:', font: {bold: true}}, {text: ' the rest'}]);

  assert.deepEqual(cell.value, {
    richText: [
      {text: 'Note:', font: {name: 'Courier New', size: 16, color: {theme: 1}, bold: true}},
      {text: ' the rest', font: {name: 'Courier New', size: 16, color: {theme: 1}}},
    ],
  });
});

test('a facet the run names wins over the cell’s', () => {
  const cell = new Cell(1, 1);
  cell.font = {name: 'Courier New', size: 16};
  cell.setRichText([{text: 'big', font: {size: 24}}]);
  assert.deepEqual(cell.value, {richText: [{text: 'big', font: {name: 'Courier New', size: 24}}]});
});

test('a cell with no font of its own passes its runs through unchanged', () => {
  // Nothing to compose: an omitted facet already falls back to the workbook default, which is
  // exactly what such a cell renders in.
  const cell = new Cell(1, 1);
  cell.setRichText([{text: 'plain'}, {text: 'bold', font: {bold: true}}]);
  assert.deepEqual(cell.value, {
    richText: [{text: 'plain'}, {text: 'bold', font: {bold: true}}],
  });
});

test('assigning value directly still writes bare runs, for a caller who wants the fallback', () => {
  const cell = new Cell(1, 1);
  cell.font = {name: 'Courier New', size: 16};
  cell.value = {richText: [{text: 'bare', font: {bold: true}}]};
  assert.deepEqual(cell.value, {richText: [{text: 'bare', font: {bold: true}}]});
});

test('text renders the cell value, and an empty cell has none', () => {
  const cell = new Cell(1, 1);
  assert.equal(cell.text, '', 'an untouched cell is empty, not "null"');
  cell.value = 42;
  assert.equal(cell.text, '42');
  cell.value = {formula: 'SUM(A1:A2)', result: 7};
  assert.equal(cell.text, '7', 'a formula reads as its cached result');
  cell.setRichText([{text: 'Note:', font: {bold: true}}, {text: ' the rest'}]);
  assert.equal(cell.text, 'Note: the rest', 'rich runs flatten in order');
});

test('text ignores the number format — the style is not the value', () => {
  const cell = new Cell(1, 1);
  cell.value = 0.5;
  cell.numFmt = '0.00%';
  assert.equal(cell.text, '0.5', 'not "50.00%": the format lives on the style');
});
