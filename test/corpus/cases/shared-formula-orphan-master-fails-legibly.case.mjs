// Cluster: formulas
//
// Real-world scenario: a caller assigns a cell a shared-formula reference — "use the same
// formula as this master cell" — but the referenced master was never given a concrete
// formula, or sits below/right of the clone. Serialization cannot resolve the clone, and
// must fail with a legible error that *names the offending cell*, not a bare throw from
// deep in the writer. The always-safe alternative is to assign each cell the same concrete
// { formula, result } instead of a shared reference; that must write and round-trip cleanly.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'shared-formula-orphan-master-fails-legibly',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 676},
  cluster: 'formulas',
  description:
    'A shared-formula clone whose master has no formula fails to serialize with an error ' +
    'that names the offending cell; assigning each cell a concrete formula is the safe ' +
    'alternative that writes and round-trips.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a shared clone with no master formula fails with an error naming the cell',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.tryWriteWorkbook({
          sheets: [{name: 'S', cells: [{ref: 'B2', sharedFormula: 'A1', result: 0}]}],
        });
        assert.strictEqual(
          ok,
          false,
          'the write must fail rather than emit a broken shared formula',
        );
        assert.match(error, /master/i, 'the error explains the missing master');
        assert.match(error, /B2/, 'the error names the offending cell');
      },
    },
    {
      name: 'assigning each cell a concrete formula writes and round-trips',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await api.roundtripFormulas({
          sheets: [
            {
              name: 'S',
              cells: [
                {ref: 'A1', value: 5},
                {ref: 'B1', formula: 'A1*2', result: 10},
                {ref: 'B2', formula: 'A1*3', result: 15},
              ],
            },
          ],
        });
        assert.strictEqual(cells.B1.formula, 'A1*2', 'first concrete formula survives');
        assert.strictEqual(cells.B2.formula, 'A1*3', 'second concrete formula survives');
      },
    },
  ],
};
