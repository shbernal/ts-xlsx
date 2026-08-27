// The layer both codecs answer to, tested on its own.
//
// `xf-style.ts` was covered only through the XML and BIFF12 readers, which meant a defect in it
// showed up as two failing codecs and no failing unit, and left the one directory under `src/` with
// source and no test beside it. The corpus case that asserts an .xlsx and its .xlsb read to the same
// model rests on exactly these three functions agreeing with themselves.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {Cell} from '../../core/cell.ts';
import {NAMED_STYLE_ID} from '../../core/internal.ts';
import {CELL_STYLE_FACETS, type Font} from '../../core/style.ts';
import {applyXfToCell, numFmtCodeFor, resolveStyleTable, type XfStyle} from './xf-style.ts';

const NO_CUSTOM: ReadonlyMap<number, string> = new Map();

test('numFmtCodeFor resolves a built-in id no file has to declare', () => {
  assert.equal(numFmtCodeFor(2, NO_CUSTOM), '0.00');
  assert.equal(numFmtCodeFor(14, NO_CUSTOM), 'mm-dd-yy');
});

test('numFmtCodeFor reads id 0 as the absence of a format, not as a format named General', () => {
  assert.equal(numFmtCodeFor(0, NO_CUSTOM), undefined);
  // A cell whose style resolves to nothing must carry no numFmt at all, or every ordinary cell in a
  // re-written file gains an explicit format it never had.
  const cell = new Cell(1, 1);
  applyXfToCell(cell, {});
  assert.equal(cell.numFmt, undefined);
});

test('numFmtCodeFor answers nothing for a reserved id it does not know, rather than guessing', () => {
  assert.equal(numFmtCodeFor(26, NO_CUSTOM), undefined, 'a gap in the reserved range');
  assert.equal(numFmtCodeFor(200, NO_CUSTOM), undefined, 'a custom id the file never declared');
  assert.equal(numFmtCodeFor(1.5, NO_CUSTOM), undefined, 'and an id that is not an id');
});

test("numFmtCodeFor lets the file's own declaration win over the built-in for the same id", () => {
  const custom = new Map([
    [164, '0.000"kg"'],
    [2, '#,##0.00 [$€-1]'],
  ]);
  assert.equal(numFmtCodeFor(164, custom), '0.000"kg"');
  assert.equal(numFmtCodeFor(2, custom), '#,##0.00 [$€-1]');
});

test('applyXfToCell writes every facet of the tuple, so a new facet cannot stop at the model', () => {
  const style: XfStyle = {
    fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}},
    numFmt: '0.00',
    font: {bold: true},
    border: {top: {style: 'thin'}},
    alignment: {horizontal: 'center'},
    protection: {locked: false},
  };
  const cell = new Cell(1, 1);
  applyXfToCell(cell, style);
  for (const facet of CELL_STYLE_FACETS) {
    assert.deepEqual(cell[facet], style[facet], `the ${facet} facet reaches the cell`);
  }
});

test('applyXfToCell leaves an absent facet absent rather than writing undefined onto the cell', () => {
  const cell = new Cell(1, 1);
  cell.font = {bold: true};
  applyXfToCell(cell, {numFmt: '0'});
  assert.deepEqual(cell.font, {bold: true}, 'a facet the xf does not name is untouched');
  assert.equal(cell.border, undefined);
  assert.equal(
    Object.hasOwn(cell.style, 'border'),
    false,
    'and is not present as an undefined key',
  );
});

test('applyXfToCell carries the two non-facet links an xf holds', () => {
  const cell = new Cell(1, 1);
  applyXfToCell(cell, {quotePrefix: true, xfId: 3});
  assert.equal(cell.quotePrefix, true);
  assert.equal(cell[NAMED_STYLE_ID], 3);
});

test('applyXfToCell on no style is a no-op, which is how an unstyled cell stays unstyled', () => {
  const cell = new Cell(1, 1);
  applyXfToCell(cell, undefined);
  assert.equal(cell.font, undefined);
  assert.equal(cell[NAMED_STYLE_ID], undefined);
});

test('resolveStyleTable layers a direct xf over the named style it links to, direct winning', () => {
  const table = resolveStyleTable({
    directXfs: [{xfId: 0, numFmt: '0.00'}],
    namedXfs: [{font: {name: 'Cambria'}, numFmt: 'General', alignment: {horizontal: 'right'}}],
    labels: [{xfId: 0, name: 'Normal', builtinId: 0}],
    fonts: [{name: 'Cambria'}],
  });
  const [xf] = table.cellXfs;
  assert.equal(xf?.numFmt, '0.00', 'the facet the direct xf sets wins');
  assert.deepEqual(xf?.font, {name: 'Cambria'}, 'and one it leaves unset falls through');
  assert.deepEqual(xf?.alignment, {horizontal: 'right'});
  assert.equal(xf?.xfId, 0, 'the link is carried through so a re-write keeps it');
});

test('resolveStyleTable leaves an xf with no link alone rather than merging style zero into it', () => {
  const table = resolveStyleTable({
    directXfs: [{numFmt: '0'}],
    namedXfs: [{font: {name: 'Cambria'}}],
    labels: [],
    fonts: [],
  });
  assert.deepEqual(table.cellXfs[0], {numFmt: '0'});
});

test('resolveStyleTable tolerates a link pointing past the named layer', () => {
  const table = resolveStyleTable({
    directXfs: [{xfId: 7, numFmt: '0'}],
    namedXfs: [],
    labels: [],
    fonts: [],
  });
  assert.deepEqual(table.cellXfs[0], {xfId: 7, numFmt: '0'});
});

test('resolveStyleTable titles the named layer by xfId, first label winning a duplicate', () => {
  const table = resolveStyleTable({
    directXfs: [],
    namedXfs: [{font: {name: 'Cambria'}}, {numFmt: '0.00'}],
    // A foreign file may name the same xfId twice, and neither reading is more right than the other.
    labels: [
      {xfId: 1, name: 'Comma', builtinId: 3},
      {xfId: 1, name: 'Comma [0]', builtinId: 6},
      {xfId: 0, name: 'Normal', builtinId: 0},
    ],
    fonts: [],
  });
  assert.deepEqual(table.namedStyles, [
    {font: {name: 'Cambria'}, name: 'Normal', builtinId: 0},
    {numFmt: '0.00', name: 'Comma', builtinId: 3},
  ]);
});

test('resolveStyleTable carries font 0 out whole as the workbook default, and omits an absent one', () => {
  const declared: Font = {name: 'Aptos Narrow', size: 11};
  const withFont = resolveStyleTable({
    directXfs: [],
    namedXfs: [],
    labels: [],
    fonts: [declared, {name: 'Consolas'}],
  });
  assert.deepEqual(withFont.defaultFont, declared);

  const withoutFont = resolveStyleTable({directXfs: [], namedXfs: [], labels: [], fonts: []});
  assert.equal(
    Object.hasOwn(withoutFont, 'defaultFont'),
    false,
    'a file with no font table declares no default, rather than an assumed Calibri',
  );
});
