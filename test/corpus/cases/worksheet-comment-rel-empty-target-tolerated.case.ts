// Cluster: address-decoding
//
// Real-world scenario: some foreign generators declare a worksheet-level comments relationship (type
// ".../comments") but leave its Target attribute empty (Target="") — they announce the relationship
// while omitting the part path. During load the reader reconciles each worksheet's relationships back
// onto the model, and for a comments relationship it dereferences the target to attach the notes. If
// that step assumes every comments relationship has a resolvable target, the empty target makes it
// read a property off `undefined` and throw, aborting the entire read of an otherwise-valid workbook.
// Reading should degrade gracefully: the file loads and the worksheets survive.
//
// This is a DISTINCT failure path from a missing VML-drawing target: here it is the comments
// relationship's own Target attribute that is blank, crashing the comment reconcile rather than the
// drawing merge. The fixture is a workbook authored with a cell note, then edited so its comments
// relationship carries Target="" while the parts are otherwise intact.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'worksheet-comment-rel-empty-target/empty-comment-rel.xlsx';

export default {
  id: 'worksheet-comment-rel-empty-target-tolerated',
  provenance: {source: 'upstream-issue'},
  cluster: 'address-decoding',
  description:
    'Loading a workbook whose worksheet declares a comments relationship with an empty Target ' +
    'attribute completes without throwing; the worksheets are recovered rather than the whole read ' +
    'aborting on the blank comment-relationship target.',

  behavior: [
    {
      name: 'loading a worksheet whose comments relationship has an empty target does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a blank comment-rel target must not abort the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheets are recovered despite the empty comment-relationship target',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(
          sheetNames,
          ['Sheet1', 'Data'],
          'both worksheets survive the tolerant read',
        );
      },
    },
  ],
} satisfies Case;
