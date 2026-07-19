// Cluster: formulas
//
// Real-world scenario: a formula cell computes a date — its cached result is stored as a date serial
// number and the cell carries a date number format (e.g. yyyy-mm-dd). When the workbook is read back,
// the formula's result value must surface as a valid Date, not an Invalid Date, so a caller reading
// the computed value gets a usable date. The format (numFmt) is what makes the numeric result a date,
// exactly as for a plain numeric-but-date-formatted cell.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'formula-date-result-reads-as-valid-date',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A formula cell whose cached result is a date serial number under a date number format reads ' +
    'back with a valid Date result (not an Invalid Date), and keeps its formula.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the formula result surfaces as a valid Date under a date number format',
      baseline: 'pass',
      async expect(api, assert) {
        const {isValidDate, resultIso} = await api.formulaDateResultReport();
        assert.strictEqual(
          isValidDate,
          true,
          `the formula result must be a valid Date; got ${resultIso}`,
        );
      },
    },
    {
      name: 'the cell keeps its formula alongside the date result',
      baseline: 'pass',
      async expect(api, assert) {
        const {keepsFormula} = await api.formulaDateResultReport();
        assert.strictEqual(keepsFormula, true, 'the cell round-trips as a formula cell');
      },
    },
  ],
};
