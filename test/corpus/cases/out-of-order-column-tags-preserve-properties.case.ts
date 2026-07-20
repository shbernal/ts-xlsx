// Cluster: styles
//
// Real-world scenario: some foreign generators (excelize, jxls-poi) emit a worksheet's <col> entries
// in non-ascending document order rather than sorted by column index. A reader that assumes ascending
// order can mis-associate widths, hidden flags, and styles with the wrong columns — or lose them. The
// reader must bind each <col> entry's properties to the column its min/max range names, regardless of
// the order the entries appear in the document. Locks foreign-generator tolerance for out-of-order
// column tags.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'out-of-order-column-tags-preserve-properties',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A worksheet whose <col> entries appear out of ascending index order (as foreign generators ' +
    'emit) still binds each column’s width and hidden flag to the correct column index on read.',

  behavior: [
    {
      name: 'each column keeps its own width despite reordered <col> tags',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {w1, w2, w3} = await api.outOfOrderColumnsReport();
        assert.strictEqual(w1, 10, 'column 1 keeps width 10');
        assert.strictEqual(w2, 20, 'column 2 keeps width 20');
        assert.strictEqual(w3, 30, 'column 3 keeps width 30');
      },
    },
    {
      name: 'a hidden column declared out of order retains its hidden flag',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hidden2} = await api.outOfOrderColumnsReport();
        assert.strictEqual(hidden2, true, 'the out-of-order hidden column stays hidden');
      },
    },
  ],
} satisfies Case;
