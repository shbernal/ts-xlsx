// Cluster: address-decoding
//
// Real-world scenario: some .xlsx files declare cell comments on a worksheet and a worksheet
// relationship pointing at a legacy VML drawing part — but the referenced VML part is not actually
// present in the package (missing, misnamed, or unreconciled). On load, the reader walks each
// worksheet relationship and, for a VML-drawing rel on a sheet that has comments, dereferences the
// drawing to merge note positioning onto the comments. When the target part is absent that lookup
// yields nothing and the merge dereference throws a TypeError, aborting the whole read even though
// the rest of the workbook is intact. Reading should degrade gracefully: the file loads, the
// worksheets survive, and the comment text (which lives in the comments part, not the VML) is kept.
//
// The fixture is a workbook authored with a cell note, then stripped of its VML drawing part while
// the worksheet relationship that references it is left in place — reproducing the dangling-target
// shape without depending on a specific application to emit it.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'load-comment-worksheet-missing-vml-drawing-target/sample.xlsx';

export default {
  id: 'load-comment-worksheet-missing-vml-drawing-target',
  provenance: {source: 'upstream-issue'},
  cluster: 'address-decoding',
  description:
    'Loading a workbook whose worksheet declares comments and a VML-drawing relationship whose ' +
    'target part is missing completes without throwing; the worksheets are recovered rather than ' +
    'the whole read aborting on the dangling drawing reference.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'loading a comment worksheet whose VML drawing target is missing does not throw',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a dangling VML-drawing target must not abort the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheets are recovered despite the missing VML drawing part',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(
          sheetNames,
          ['Sheet1', 'Data'],
          'both worksheets survive the tolerant read',
        );
      },
    },
  ],
};
