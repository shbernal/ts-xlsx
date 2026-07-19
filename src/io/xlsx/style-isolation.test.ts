// Copy-on-write style isolation.
//
// On disk, identically-formatted cells deduplicate to one shared style record, so a loaded
// workbook can hand several cells the same style. Mutating one cell's facet must change ONLY
// that cell — never a sibling that happened to share the record. The rewrite gets this by
// construction: each cell owns independent facet fields and every setter REPLACES the field
// (the facet types are `readonly`, so a shared record cannot be edited in place). These tests
// hard-lock that guarantee — legacy bled here, and a future refactor that reintroduced in-place
// mutation would silently pass the corpus (its baseline is the legacy bleed), so the lock lives
// here in src.

import assert from 'node:assert/strict';
import {test} from 'node:test';
import type {Fill} from '../../core/style.ts';
import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function roundtrip(workbook: Workbook): Workbook {
  return readXlsx(writeXlsx(workbook));
}

const fgOf = (fill: Fill | undefined): string | undefined =>
  fill?.type === 'pattern' ? fill.fgColor?.argb : undefined;

test('replacing one loaded cell fill leaves a style-sharing sibling untouched, in memory and on disk', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}} as const;
  s.getCell('A1').value = 'a';
  s.getCell('B1').value = 'b';
  s.getCell('A1').fill = fill;
  s.getCell('B1').fill = fill; // one shared style index on disk

  const loadedWb = roundtrip(wb);
  const loaded = loadedWb.getWorksheet('S');
  assert.ok(loaded);
  loaded.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  assert.equal(
    fgOf(loaded.getCell('A1').fill),
    'FF00FF00',
    'the edited cell reflects the new fill',
  );
  assert.equal(
    fgOf(loaded.getCell('B1').fill),
    'FFFF0000',
    'the sibling keeps its original fill in memory',
  );

  const back = roundtrip(loadedWb).getWorksheet('S');
  assert.equal(fgOf(back?.getCell('A1').fill), 'FF00FF00', 'the edit persists to disk');
  assert.equal(fgOf(back?.getCell('B1').fill), 'FFFF0000', 'only the edited cell changed on disk');
});

test('spread-reassigning one loaded cell font member does not bleed into a shared sibling', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  const font = {name: 'Arial', size: 12, color: {argb: 'FF000000'}} as const;
  s.getCell('A1').value = 'a';
  s.getCell('A1').font = font;
  s.getCell('B1').value = 'b';
  s.getCell('B1').font = font; // one shared style index on disk

  const loaded = roundtrip(wb).getWorksheet('S');
  assert.ok(loaded);
  const a1 = loaded.getCell('A1');
  a1.font = {...a1.font, color: {argb: 'FFFF0000'}};
  assert.equal(
    loaded.getCell('A1').font?.color?.argb,
    'FFFF0000',
    'the edited cell reflects the new color',
  );
  assert.equal(
    loaded.getCell('B1').font?.color?.argb,
    'FF000000',
    'the sibling keeps its original color',
  );
});

test('assigning the same base font object to two cells then mutating one isolates the sibling', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  const base = {name: 'Arial', size: 11} as const;
  s.getCell('A1').value = 'YES';
  s.getCell('A2').value = 'NO';
  s.getCell('A1').font = base;
  s.getCell('A2').font = base; // the SAME object assigned to both cells — the aliasing trap
  const a1 = s.getCell('A1');
  a1.font = {...a1.font, color: {argb: 'FF00FF00'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').font?.color?.argb,
    'FF00FF00',
    'the targeted cell carries the new color',
  );
  assert.equal(
    back?.getCell('A2').font?.color,
    undefined,
    'the sibling given the same base keeps no color',
  );
});

test('bordering one loaded cell that shares a style record borders only that cell', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  for (const r of [1, 2, 3]) {
    const c = s.getCell(`A${r}`);
    c.value = 'x';
    c.font = {bold: true}; // identical → one shared style index on disk
  }

  const loadedWb = roundtrip(wb);
  const loaded = loadedWb.getWorksheet('S');
  assert.ok(loaded);
  loaded.getCell('A1').border = {
    top: {style: 'thin'},
    left: {style: 'thin'},
    bottom: {style: 'thin'},
    right: {style: 'thin'},
  };

  const back = roundtrip(loadedWb).getWorksheet('S');
  const hasTop = (ref: string): boolean => !!back?.getCell(ref).border?.top?.style;
  assert.equal(hasTop('A1'), true, 'the targeted cell gains the border');
  assert.equal(hasTop('A2'), false, 'a sibling gains no border');
  assert.equal(hasTop('A3'), false, 'a sibling gains no border');
});

test('setting one facet on a loaded cell keeps the sibling and never drops the cell’s other facets', () => {
  // Both cells share a numFmt-only style; A1 gains an alignment. The sibling must keep just the
  // numFmt, and A1 must keep BOTH its numFmt and the new alignment (facets compose, not replace).
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'a';
  s.getCell('B1').value = 'b';
  s.getCell('A1').numFmt = '0.00';
  s.getCell('B1').numFmt = '0.00'; // one shared style index on disk

  const loadedWb = roundtrip(wb);
  const loaded = loadedWb.getWorksheet('S');
  assert.ok(loaded);
  loaded.getCell('A1').alignment = {horizontal: 'center'};

  const back = roundtrip(loadedWb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').alignment?.horizontal,
    'center',
    'the edited cell gains the alignment',
  );
  assert.equal(back?.getCell('A1').numFmt, '0.00', 'the edited cell keeps its number format');
  assert.equal(back?.getCell('B1').alignment, undefined, 'the sibling gains no alignment');
  assert.equal(back?.getCell('B1').numFmt, '0.00', 'the sibling keeps its number format');
});
