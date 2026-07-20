// Cluster: comments
//
// Real-world scenario: a workbook from a foreign generator (easyexcel) has cell comments whose
// text runs are empty — every comment element is a bare, text-less note (with its VML drawing
// wiring). Reading such a file must complete without error and surface each comment as an empty
// note, consistently across every affected cell — an empty comment is a valid comment, not a
// parse failure and not a missing note.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'empty-comments-read-as-blank-notes/sample.xlsx';

export default {
  id: 'empty-comments-read-as-blank-notes',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2155},
  cluster: 'comments',
  description:
    'A workbook whose comments have no text runs reads without error and surfaces each as an ' +
    'empty note, consistently — an empty comment is a valid, blank note rather than a crash.',

  behavior: [
    {
      name: 'a workbook with empty comments reads without error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `empty comments must not crash the read; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'an empty comment surfaces as an empty note',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {A1} = await api.readFixtureCells(FIXTURE, ['A1']);
        assert.notStrictEqual(A1.note, undefined, 'the commented cell exposes a note');
        assert.strictEqual(A1.note, '', 'the empty comment reads as an empty-string note');
      },
    },
  ],
} satisfies Case;
