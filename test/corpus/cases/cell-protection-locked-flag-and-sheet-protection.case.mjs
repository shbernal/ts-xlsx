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

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'cell-protection-locked-flag-and-sheet-protection',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Per-cell protection round-trips the meaningful flag (an explicitly unlocked cell survives as ' +
    'locked=false, carried via applyProtection in the style record), and protecting the worksheet ' +
    'emits the <sheetProtection> element that makes locked flags enforceable.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an explicitly unlocked cell round-trips as locked=false while a default-protection sibling does not report unlocked',
      baseline: 'pass',
      async expect(api, assert) {
        const {readBack} = await api.authorCellProtection([
          {ref: 'A1', value: 'default'},
          {ref: 'A2', value: 'editable', protection: {locked: false}},
        ]);
        assert.ok(readBack.A2 && readBack.A2.locked === false, `the unlocked cell must round-trip as locked=false; got ${JSON.stringify(readBack.A2)}`);
        assert.ok(
          !readBack.A1 || readBack.A1.locked !== false,
          `a default cell must not come back as explicitly unlocked; got ${JSON.stringify(readBack.A1)}`
        );
      },
    },
    {
      name: 'setting a cell unlocked carries the flag into the style record (applyProtection + <protection> in cellXfs)',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasApplyProtection} = await api.authorCellProtection([
          {ref: 'A2', value: 'editable', protection: {locked: false}},
        ]);
        assert.ok(hasApplyProtection, 'the written style record must mark applyProtection with a <protection> child, not drop the flag');
      },
    },
    {
      name: 'protecting the worksheet emits a <sheetProtection> element — the thing that makes locked flags enforceable',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetProtection} = await api.authorCellProtection(
          [{ref: 'A1', value: 'x'}],
          {password: 'secret', options: {selectLockedCells: true}}
        );
        assert.ok(sheetProtection, 'protecting the sheet must emit a <sheetProtection> element');
        assert.ok(/sheet="1"/.test(sheetProtection), `sheetProtection must enable sheet-level locking; got ${sheetProtection}`);
      },
    },
  ],
};
