// Cluster: styles
//
// Real-world scenario: a user reads a template, then inserts rows in a style-inheritance mode (copy
// the style of an adjacent row) rather than the no-style mode. A reported failure was that the
// inserted cells' style objects were frozen internally, so any later style assignment — setting
// numFmt to a currency pattern, toggling font bold — threw "Cannot add property numFmt, object is not
// extensible", forcing the user back to the no-style mode. Inserting a row and then styling its cells
// must work regardless of the style-inheritance mode.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'inserted-row-cells-stay-mutable-with-style-inheritance',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'After inserting a row with a style-inheritance mode, assigning a numFmt and mutating a font ' +
    'property on the inserted cells succeeds (no "object is not extensible" throw) and the numFmt is ' +
    'applied — behaving like the no-style insert mode.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'styling a cell of an inheritance-inserted row does not throw a non-extensible error',
      baseline: 'pass',
      async expect(api, assert) {
        const {error} = await api.insertRowThenStyle('i');
        assert.strictEqual(error, null, `styling an inherited-style cell must not throw; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the assigned numFmt is applied to the inserted cell',
      baseline: 'pass',
      async expect(api, assert) {
        const {numFmt} = await api.insertRowThenStyle('i');
        assert.strictEqual(numFmt, '$#,##0.00;[Red]-$#,##0.00', 'the numFmt is set on the inserted cell');
      },
    },
    {
      name: 'the no-style insert mode also leaves cells mutable (consistency)',
      baseline: 'pass',
      async expect(api, assert) {
        const {error} = await api.insertRowThenStyle('n');
        assert.strictEqual(error, null, 'the no-style mode leaves cells mutable too');
      },
    },
  ],
};
