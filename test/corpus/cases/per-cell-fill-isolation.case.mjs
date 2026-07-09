// Cluster: styles
//
// Real-world scenario: a script reads a workbook and paints a background fill onto
// one specific cell (e.g. to flag a dirty/edited value). Only that cell must end up
// filled. A recurring complaint is that setting a fill "also fills other cells" —
// style bleeding across a row, column, or the whole sheet. A fill assigned to one
// cell must stay local to that cell through a write→read round-trip; untouched
// cells must come back with no fill.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const RED = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'painted', fill: RED},
        {ref: 'B2', value: 'untouched'},
        {ref: 'A2', value: 'same-column'},
        {ref: 'C1', value: 'same-row'},
      ],
    },
  ],
};

export default {
  id: 'per-cell-fill-isolation',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 791},
  cluster: 'styles',
  description:
    'A background fill assigned to a single cell stays local to that cell across a ' +
    'write→read round-trip; it does not bleed onto other cells in the same row, the ' +
    'same column, or elsewhere on the sheet.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a fill set on A1 is observable on A1',
      baseline: 'pass',
      async expect(api, assert) {
        const {A1} = (await api.roundtripWorkbook(SPEC)).sheets.S.cells;
        assert.ok(A1.fill && A1.fill.type === 'pattern', 'A1 should keep its fill');
        assert.strictEqual(A1.fill.fgColor.argb, 'FFFF0000');
      },
    },
    {
      name: 'the fill does not bleed onto an untouched cell (B2)',
      baseline: 'pass',
      async expect(api, assert) {
        const {B2} = (await api.roundtripWorkbook(SPEC)).sheets.S.cells;
        assert.ok(!B2.fill || B2.fill.type === undefined, 'B2 must have no fill');
      },
    },
    {
      name: 'the fill does not bleed along the same column (A2) or row (C1)',
      baseline: 'pass',
      async expect(api, assert) {
        const {A2, C1} = (await api.roundtripWorkbook(SPEC)).sheets.S.cells;
        assert.ok(!A2.fill || A2.fill.type === undefined, 'A2 (same column) must have no fill');
        assert.ok(!C1.fill || C1.fill.type === undefined, 'C1 (same row) must have no fill');
      },
    },
  ],
};
