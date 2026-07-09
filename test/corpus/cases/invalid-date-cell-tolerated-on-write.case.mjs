// Cluster: dates
//
// Real-world scenario: a data pipeline maps a column to Date values, and one row's
// source is null or malformed, producing `new Date(NaN)` (an Invalid Date). Writing
// the workbook must not throw an unhandled error out of the write call, and — most
// importantly — must not lose the other, valid cells in that sheet. A single bad
// value should degrade locally, never take down the whole export.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

const SPEC = {
  sheets: [{
    name: 'Data',
    cells: [
      {ref: 'A1', value: {invalidDate: true}},
      {ref: 'B1', value: 'still here'},
      {ref: 'C1', value: 42},
    ],
  }],
};

export default {
  id: 'invalid-date-cell-tolerated-on-write',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 449},
  cluster: 'dates',
  description:
    'A cell holding an Invalid Date (new Date(NaN)) must not make workbook write ' +
    'throw, and must not cause sibling cells to be lost.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'writing a workbook with an invalid-date cell does not throw',
      baseline: 'pass',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(SPEC);
        assert.ok(result.ok, `write failed (${result.phase}): ${result.error}`);
      },
    },
    {
      name: 'an invalid-date cell does not drop its sibling cells',
      baseline: 'pass',
      async expect(api, assert) {
        const {survivingCells} = await api.tryWriteWorkbook(SPEC);
        assert.ok(survivingCells.Data.includes('B1'), 'sibling string cell B1 was lost');
        assert.ok(survivingCells.Data.includes('C1'), 'sibling number cell C1 was lost');
      },
    },
  ],
};
