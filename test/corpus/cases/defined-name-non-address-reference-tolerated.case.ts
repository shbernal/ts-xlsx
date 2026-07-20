// Cluster: address-decoding
//
// Real-world scenario: some workbooks declare defined names (named ranges) whose stored formula is
// not a plain cell/range address — a constant value, an error reference like #REF!, or a name into
// an external workbook. A reader that decodes every defined name as an address during reconcile can
// throw on the degenerate one ("Cannot read property 'match' of undefined") and abort the entire
// load. A workbook must stay readable even when one of its defined names does not resolve to a normal
// in-sheet address: the offending name is skipped, the rest of the workbook — other worksheets, cell
// values, and well-formed defined names — comes through intact.
//
// The fixture is a workbook whose workbook.xml declares a mix of names: a valid one
// (`GoodName` → Sheet1!$A$1), a constant (`42`), an error ref (`#REF!`), and an external
// reference. It reproduces the shape of files seen in the wild without depending on any one
// application to author the degenerate entries.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'defined-name-non-address-reference-tolerated/sample.xlsx';

export default {
  id: 'defined-name-non-address-reference-tolerated',
  provenance: {source: 'upstream-issue'},
  cluster: 'address-decoding',
  description:
    'Reading a workbook whose defined names include non-address references (a constant, an error ' +
    'reference, an external-workbook name) completes without throwing; the worksheets and their ' +
    'values are recovered and a well-formed defined name alongside the degenerate ones survives.',

  behavior: [
    {
      name: 'a workbook with non-address defined names reads without throwing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `the load must not abort on a degenerate defined name; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'every worksheet is recovered intact',
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
    {
      name: 'a valid defined name survives even though its degenerate siblings do not decode',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {names} = await api.readFixtureDefinedNames(FIXTURE);
        assert.deepStrictEqual(
          names.GoodName,
          ['Sheet1!$A$1'],
          'the well-formed defined name is read back rather than dropped alongside the degenerate ones',
        );
      },
    },
  ],
} satisfies Case;
