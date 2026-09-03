// Cluster: styles
//
// Real-world scenario: a sheet has content in B2 and C2, then those two cells are merged. C2 is still
// a materialised cell carrying its own formatting: merging hides it from the user, it does not delete
// it, and the writer has to know it is there.
//
// A range over the merged block enumerated positions and fetched each by address, but fetching a
// covered address resolves it to the merge master. So the range saw the master twice and the covered
// cell never: `cells` reported a duplicate, and `clearStyle` cleared the master twice while leaving
// the covered cell exactly as styled as it found it, which is the opposite of what it documents.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-range-over-a-merge-sees-each-stored-cell-once',
  provenance: {source: 'grid-edit-audit'},
  cluster: 'styles',
  description:
    'A range enumerates the cells actually stored under it, each exactly once, rather than ' +
    'resolving covered addresses to the merge master: so a merged block reports both its cells and ' +
    'clearStyle clears both.',

  behavior: [
    {
      name: 'a range over a merge reports the master and the covered cell, not the master twice',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.rangeOverMergeReport().addresses, ['B2', 'C2']);
      },
    },
    {
      name: 'clearStyle over a merge clears the covered cell as well as the master',
      expect(api: CorpusApi, assert: Assert) {
        const {numFmtsAfterClear} = api.rangeOverMergeReport();
        assert.equal(numFmtsAfterClear['B2'], null, 'the master is cleared');
        assert.equal(numFmtsAfterClear['C2'], null, 'and so is the cell the merge covers');
      },
    },
  ],
} satisfies Case;
