// Cluster: core-model
//
// Real-world scenario: a horizontal merge (A1:B1) has its value in the master cell (A1); the other
// cells of the range are merged children. Reading the displayed text of a merged child cell must not
// throw and must return the same text as the master — a consumer iterating the cells of a merged
// range should see consistent text at every cell, not an exception on the child cells. Locks the
// merged-child text behavior against the reported throw.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'merged-child-cell-text-mirrors-master',
  provenance: {source: 'upstream-issue'},
  cluster: 'core-model',
  description:
    'Reading the display text of a merged child (non-master) cell returns the master cell’s text ' +
    'and does not throw, so text reads are consistent across every cell of a merged range.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the merge master cell reports its own text',
      baseline: 'pass',
      async expect(api, assert) {
        const {masterText} = await api.mergedCellDisplayTextReport();
        assert.strictEqual(masterText, 'Group', 'the master reports its value as text');
      },
    },
    {
      name: 'a merged child cell mirrors the master text without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {childText, childThrew} = await api.mergedCellDisplayTextReport();
        assert.strictEqual(childThrew, false, 'reading a merged child cell’s text must not throw');
        assert.strictEqual(childText, 'Group', 'the child text mirrors the master');
      },
    },
  ],
};
