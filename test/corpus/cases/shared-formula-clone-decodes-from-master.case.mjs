// Cluster: formulas
//
// Real-world scenario: spreadsheets store a formula filled down a column once, on a
// "master" cell, and mark every other cell in the range as a bare clone that only points
// back at the master by a shared index — the clone carries no formula text of its own.
// Reading such a file, each clone must resolve to a concrete formula: the master's formula
// address-translated to the clone's position (a master `A1*2` at B1, shared down, means B2
// is `A2*2`, B3 is `A3*2`). A reader that left clones empty would lose every dragged-down
// formula in the sheet. The master keeps its own formula and every cell keeps its cached
// result.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A2', value: 2},
        {ref: 'A3', value: 3},
        {ref: 'B1', formula: 'A1*2', result: 2},
        {ref: 'B2', sharedFormula: 'B1', result: 4},
        {ref: 'B3', sharedFormula: 'B1', result: 6},
      ],
    },
  ],
};

export default {
  id: 'shared-formula-clone-decodes-from-master',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 260},
  cluster: 'formulas',
  description:
    'A shared-formula clone (a cell that references a master by shared index rather than ' +
    'carrying formula text) reads back a concrete formula that is the master’s formula ' +
    'translated to the clone’s address, with its cached result preserved.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the master cell keeps its own formula and result',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await api.roundtripFormulas(SPEC);
        assert.strictEqual(cells.B1.formula, 'A1*2', 'master formula survives');
        assert.strictEqual(cells.B1.result, 2, 'master result survives');
      },
    },
    {
      name: 'a shared clone resolves to the master formula translated to its own address',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await api.roundtripFormulas(SPEC);
        assert.strictEqual(cells.B2.formula, 'A2*2', 'clone one row down is A2*2, not empty');
        assert.strictEqual(cells.B3.formula, 'A3*2', 'clone two rows down is A3*2');
        assert.strictEqual(cells.B2.sharedFormula, 'B1', 'the clone still records its master reference');
      },
    },
    {
      name: 'a shared clone preserves its cached result',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await api.roundtripFormulas(SPEC);
        assert.strictEqual(cells.B2.result, 4, 'clone result survives');
        assert.strictEqual(cells.B3.result, 6, 'clone result survives');
      },
    },
  ],
};
