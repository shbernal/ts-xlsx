// Cluster: core-model
//
// Real-world scenario: a worksheet has rows carrying metadata beyond cell values — a cell note
// (comment) and an outline (grouping) level. When a row is inserted above them, that metadata must
// travel with its logical row: the note stays attached to the same cell content at its shifted
// position, and the outline level stays on the row it grouped. The observed bugs: inserting a row
// drops the cell note entirely, and the outline level stays pinned to the old absolute row index
// rather than following its row — so after an insert the grouping lands on the wrong row and the note
// is gone. Cell values themselves shift correctly, which isolates the defect to the row metadata.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'row-insert-preserves-note-and-outline-level',
  provenance: {source: 'upstream-issue'},
  cluster: 'core-model',
  description:
    'Inserting a row above rows that carry a cell note and an outline level shifts those rows and ' +
    'keeps their metadata attached to the correct logical row: the note survives at the shifted ' +
    'cell and the outline level follows its row rather than staying at the old absolute index.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'cell values shift down correctly on insert (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {dataShifted} = await api.rowInsertPreservesNoteAndOutline();
        assert.strictEqual(dataShifted, true, 'the pre-existing rows move down by one');
      },
    },
    {
      name: 'a cell note survives a row insert above it',
      baseline: 'pass',
      async expect(api, assert) {
        const {noteFollowsRow} = await api.rowInsertPreservesNoteAndOutline();
        assert.strictEqual(
          noteFollowsRow,
          true,
          'the note must remain attached to its cell at the shifted position, not be dropped',
        );
      },
    },
    {
      name: 'an outline level follows its row through a row insert',
      baseline: 'pass',
      async expect(api, assert) {
        const {outlineFollowsRow} = await api.rowInsertPreservesNoteAndOutline();
        assert.strictEqual(
          outlineFollowsRow,
          true,
          'the outline level must move with its logical row, not stay at the old absolute row index',
        );
      },
    },
  ],
};
