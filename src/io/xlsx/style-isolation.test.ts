// Copy-on-write style isolation. On disk, identically-formatted cells deduplicate to one shared style
// record, so a loaded workbook can hand several cells the same style. The corpus locks that editing one
// such cell leaves its siblings alone; this adds that the edited cell keeps the facets it already had.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {roundtrip} from './package.test-support.ts';

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
