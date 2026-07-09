// Cluster: core-model
//
// Real-world scenario: a user duplicates a worksheet by reading one sheet's `model` object and
// assigning it onto a freshly added worksheet (the idiomatic "clone a sheet" pattern). The source
// sheet has merged cell ranges. After the assignment the destination has all the values, columns,
// and styles — but none of the merged ranges: the merges silently vanish. The root cause is an
// asymmetry in the worksheet model's serialize/deserialize contract — the exported model exposes
// merged ranges under one property while the importer reads a different one — so merge data does not
// survive a model export/import round-trip. Users worked around it for years by manually
// re-applying every merge after the copy.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'worksheet-model-preserves-merged-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'core-model',
  description:
    'Copying a worksheet via its model export/import (dst.model = {...src.model, name}) reproduces ' +
    'the source sheet\'s merged ranges on the destination — model round-trip is lossless for merges.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the exported source model carries the merged ranges (precondition)',
      baseline: 'pass',
      expect(api, assert) {
        const {srcMerges} = api.copyWorksheetModel({merges: ['A1:C1']});
        assert.deepStrictEqual(srcMerges, ['A1:C1'], 'the source model exposes its merge ranges');
      },
    },
    {
      name: 'assigning that model onto another sheet reproduces the merged ranges',
      baseline: 'fail',
      expect(api, assert) {
        const {dstMerges, error} = api.copyWorksheetModel({merges: ['A1:C1']});
        assert.strictEqual(error, null, 'the model copy must not throw');
        assert.deepStrictEqual(
          dstMerges,
          ['A1:C1'],
          'the destination sheet must carry the same merges the source model exported'
        );
      },
    },
  ],
};
