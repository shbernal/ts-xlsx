// Cluster: core-model
//
// Real-world scenario: in Excel a user selects a *block* of cells at once, opens Data
// Validation, and applies one rule (e.g. "whole number between 0 and 9"). Excel stores
// that as a single validation whose `sqref` covers the whole selected range rather than
// one entry per cell. When the file is read back, every cell in that range must report
// the validation — a program that inspects one cell of the block to learn its rule must
// not find the rule missing just because it was authored across a multi-cell selection.
//
// Fixture `book.xlsx` was authored in Excel: A1:A3 share a "whole between 0 and 9"
// validation applied to the multi-cell selection, and B2 has its own separate rule.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'data-validation-multi-cell-selection/book.xlsx';

export default {
  id: 'data-validation-multi-cell-selection',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 317},
  cluster: 'core-model',
  description:
    'A data validation applied in Excel to a multi-cell selection is read back on every ' +
    'cell of the selected range, not only on the first — the range-form (sqref spanning ' +
    'multiple cells) is expanded per cell on read.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every cell of the multi-cell selection reports the shared validation',
      baseline: 'pass',
      async expect(api, assert) {
        const {cells} = await api.readFixtureValidations(FIXTURE);
        for (const ref of ['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3']) {
          const dv = cells[ref];
          assert.ok(dv, `${ref} should carry the range's validation`);
          assert.strictEqual(dv.type, 'whole', `${ref} type`);
          assert.strictEqual(dv.operator, 'between', `${ref} operator`);
          assert.deepStrictEqual(dv.formulae, [0, 9], `${ref} bounds`);
        }
      },
    },
    {
      name: 'a separately-validated cell keeps its own distinct rule',
      baseline: 'pass',
      async expect(api, assert) {
        const {cells} = await api.readFixtureValidations(FIXTURE);
        const dv = cells['Sheet1!B2'];
        assert.ok(dv, 'B2 should carry its own validation');
        assert.strictEqual(dv.operator, 'equal', 'B2 has its own operator, not the range rule');
        assert.deepStrictEqual(dv.formulae, [8], 'B2 keeps its own bound');
      },
    },
  ],
};
