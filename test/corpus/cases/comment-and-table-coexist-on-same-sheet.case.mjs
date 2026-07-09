// Cluster: tables
//
// Real-world scenario: a worksheet contains both a table and at least one cell comment
// (note). Each feature adds its own worksheet relationships and parts — a comment brings a
// comments part plus a legacy VML drawing, a table brings a table part — and if their
// relationship wiring collides the written file is corrupt: Excel repairs it on open and
// can blank the affected cells. A sheet with both a table and a comment must produce a
// valid, internally-consistent package where both survive.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [{ref: 'E5', value: 'note here', note: 'a comment'}],
      tables: [{name: 'T', ref: 'A1', headers: ['H1', 'H2'], rows: [['a', 1], ['b', 2]]}],
    },
  ],
};

export default {
  id: 'comment-and-table-coexist-on-same-sheet',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1429},
  cluster: 'tables',
  description:
    'A worksheet carrying both a table and a cell comment writes a valid package — the ' +
    'comments part, its VML drawing, and the table part coexist with unique worksheet ' +
    'relationship ids — and both the table data and the comment survive a round-trip.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the package carries both a comments part and a table part with unique worksheet rels',
      baseline: 'pass',
      async expect(api, assert) {
        const {packageParts, consistency} = await api.inspectPackage(SPEC);
        assert.ok(packageParts.hasCommentsPart, 'a comments part is written');
        assert.ok(packageParts.hasVmlDrawingPart, 'the comment’s VML drawing is written');
        assert.ok(packageParts.hasTablePart, 'a table part is written');
        assert.ok(
          consistency.worksheetRelIdsUnique,
          'the worksheet relationship ids must be unique across comment, VML, and table'
        );
      },
    },
    {
      name: 'the table data and the comment both survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const e5 = model.sheets.S.cells.E5;
        assert.strictEqual(e5.value, 'note here', 'the commented cell keeps its value (sheet not blanked)');
        assert.strictEqual(e5.note, 'a comment', 'the comment survives alongside the table');
      },
    },
  ],
};
