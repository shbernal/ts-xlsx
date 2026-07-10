// Cluster: comments
//
// Real-world scenario: a user attaches a note (cell comment) to a cell, then decides it should not
// appear in the exported file. There is no first-class way to remove a note: assigning an empty note
// object (the common workaround) persists an *empty* comment — the cell still shows a comment marker
// in spreadsheet apps and the package still carries the comment part and its VML drawing. What the
// user needs is a genuine removal: after clearing, the cell reports no note and the written package
// contains no comment or VML artifact for it, while other cells' notes and values are untouched.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'removing-a-cell-note-leaves-no-comment-artifact',
  provenance: {source: 'upstream-issue'},
  cluster: 'comments',
  description:
    'Clearing a cell note produces a genuinely note-free cell: the round-tripped package carries no ' +
    'comment part or VML drawing for it, while a neighboring cell that kept its note is unaffected. ' +
    'A workbook whose cell never had a note emits no comment part at all.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook whose cell never carried a note emits no comment part (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {cleanHasCommentPart} = await api.removeCellNoteReport();
        assert.strictEqual(cleanHasCommentPart, false, 'no note anywhere means no comment part is written');
      },
    },
    {
      name: 'clearing the only relevant note removes its comment part from the package',
      baseline: 'fail',
      async expect(api, assert) {
        const {commentPartPresent, readNoteAfter} = await api.removeCellNoteReport();
        assert.strictEqual(
          commentPartPresent,
          false,
          `a cleared note must leave no comment part; the cell still reads back a note=${JSON.stringify(readNoteAfter)}`
        );
      },
    },
    {
      name: 'clearing one note leaves a different cell’s note intact',
      baseline: 'pass',
      async expect(api, assert) {
        const {neighborNoteIntact} = await api.removeCellNoteReport();
        assert.strictEqual(neighborNoteIntact, true, 'removing one note must not disturb another cell’s note');
      },
    },
  ],
};
