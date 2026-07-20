// Cluster: styles
//
// Real-world scenario: an author wants some cells editable and the rest read-only. In OOXML the
// per-cell "locked" flag defaults to TRUE — so marking a cell locked=true is a no-op relative to the
// default, and the only per-cell state that actually carries information is locked=FALSE (an
// explicitly *unlocked* cell). Crucially, the locked flag does nothing on its own: it is enforced
// only when the worksheet itself is protected (a <sheetProtection> element). A common confusion is
// setting locked=true and expecting cells to become read-only without ever protecting the sheet.
//
// The durable guarantees: an unlocked cell must survive a round-trip (so the author's editable
// regions are preserved), the unlocked flag must actually be carried in the cell's style record
// (applyProtection + <protection> in cellXfs, not silently dropped), and protecting the sheet must
// emit the <sheetProtection> element that makes the locked flags enforceable by a consumer.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'cell-protection-locked-flag-and-sheet-protection',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Per-cell protection round-trips the meaningful flag (an explicitly unlocked cell survives as ' +
    'locked=false, carried via applyProtection in the style record), and protecting the worksheet ' +
    'emits the <sheetProtection> element that makes locked flags enforceable.',

  behavior: [
    {
      name: 'an explicitly unlocked cell round-trips as locked=false while a default-protection sibling does not report unlocked',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {readBack} = await api.authorCellProtection([
          {ref: 'A1', value: 'default'},
          {ref: 'A2', value: 'editable', protection: {locked: false}},
        ]);
        assert.ok(
          readBack.A2 && readBack.A2.locked === false,
          `the unlocked cell must round-trip as locked=false; got ${JSON.stringify(readBack.A2)}`,
        );
        assert.ok(
          readBack.A1?.locked !== false,
          `a default cell must not come back as explicitly unlocked; got ${JSON.stringify(readBack.A1)}`,
        );
      },
    },
    {
      name: 'setting a cell unlocked carries the flag into the style record (applyProtection + <protection> in cellXfs)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasApplyProtection} = await api.authorCellProtection([
          {ref: 'A2', value: 'editable', protection: {locked: false}},
        ]);
        assert.ok(
          hasApplyProtection,
          'the written style record must mark applyProtection with a <protection> child, not drop the flag',
        );
      },
    },
    {
      name: 'protecting the worksheet emits a <sheetProtection> element — the thing that makes locked flags enforceable',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetProtection} = await api.authorCellProtection([{ref: 'A1', value: 'x'}], {
          password: 'secret',
          options: {selectLockedCells: true},
        });
        assert.ok(sheetProtection, 'protecting the sheet must emit a <sheetProtection> element');
        assert.ok(
          /sheet="1"/.test(sheetProtection),
          `sheetProtection must enable sheet-level locking; got ${sheetProtection}`,
        );
      },
    },
    {
      // Unlocking a whole COLUMN or ROW in one call (rather than touching each cell) must carry the
      // unlocked flag to every cell of that band, exactly as a per-cell override would — the band is
      // just an ergonomic shorthand for "these cells are editable once the sheet is protected".
      name: 'setting a whole column unlocked carries locked=false to its cells; an off-band cell stays default-locked',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {readBack} = await api.authorCellProtection(
          [
            {ref: 'A1', value: 'a'},
            {ref: 'B1', value: 'b'},
          ],
          {password: 'pw'},
          {columns: [{index: 1, protection: {locked: false}}]},
        );
        assert.ok(
          readBack.A1 && readBack.A1.locked === false,
          `a cell in the unlocked column must round-trip locked=false; got ${JSON.stringify(readBack.A1)}`,
        );
        assert.ok(
          readBack.B1?.locked !== false,
          `a cell outside the unlocked column must not be explicitly unlocked; got ${JSON.stringify(readBack.B1)}`,
        );
      },
    },
    {
      name: 'setting a whole row unlocked carries locked=false to its cells',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {readBack} = await api.authorCellProtection(
          [{ref: 'A3', value: 'c'}],
          {password: 'pw'},
          {rows: [{index: 3, protection: {locked: false}}]},
        );
        assert.ok(
          readBack.A3 && readBack.A3.locked === false,
          `a cell in the unlocked row must round-trip locked=false; got ${JSON.stringify(readBack.A3)}`,
        );
      },
    },
  ],
} satisfies Case;
