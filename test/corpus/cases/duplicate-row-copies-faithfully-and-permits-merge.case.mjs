// Cluster: tables
//
// Real-world scenario: a worksheet's row-duplication operation copies a source row's values to a new
// row. Two reported concerns: (1) invoked with default arguments (no count, no insert flag), it must
// produce exactly one faithful copy of the source values after the source, not empty/NaN/garbage
// cells; and (2) duplicating a row that carries no merge must not fabricate a phantom merge on the
// new row, so a subsequent explicit merge of a range on the duplicated row succeeds rather than
// failing with an "already merged" error.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'duplicate-row-copies-faithfully-and-permits-merge',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Duplicating a populated row yields exactly one faithful copy of its values, and merging a range ' +
    'on the duplicated (previously unmerged) row succeeds without a phantom "already merged" error.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the duplicated row is a faithful copy of the source values',
      baseline: 'pass',
      async expect(api, assert) {
        const {row1, row2, dupError} = await api.duplicateRowReport();
        assert.strictEqual(
          dupError,
          null,
          `duplication must not throw; got ${JSON.stringify(dupError)}`,
        );
        assert.deepStrictEqual(row1, ['a', 'b', 'c'], 'the source row keeps its values');
        assert.deepStrictEqual(
          row2,
          ['a', 'b', 'c'],
          'the duplicated row copies the values (not empty/NaN)',
        );
      },
    },
    {
      name: 'merging a range on the duplicated row succeeds without an "already merged" error',
      baseline: 'pass',
      async expect(api, assert) {
        const {mergeError, merges} = await api.duplicateRowReport();
        assert.strictEqual(
          mergeError,
          null,
          `merging the duplicated row must not throw; got ${JSON.stringify(mergeError)}`,
        );
        assert.ok(
          merges.includes('A2:C2'),
          `the intended merge is present; got ${JSON.stringify(merges)}`,
        );
      },
    },
  ],
};
