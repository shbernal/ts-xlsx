// Cluster: xlsx-io
//
// Real-world scenario: a worksheet produced by a foreign (non-Excel) generator omits the
// optional <sheetFormatPr> element (default row height / outline properties) that Excel always
// writes. A reader that assumes the element is present crashes dereferencing outline properties
// on a null format object. A missing <sheetFormatPr> must be treated as default/empty format
// properties, and the file must read without throwing — its sheets discoverable and its cells
// readable.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'foreign-sheet-without-sheetformatpr-reads/sample.xlsx';

export default {
  id: 'foreign-sheet-without-sheetformatpr-reads',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1185},
  cluster: 'xlsx-io',
  description:
    'A worksheet from a foreign generator that omits the optional <sheetFormatPr> element reads ' +
    'without crashing — a missing format-properties element is treated as defaults, and the ' +
    'sheet remains discoverable by name.',

  behavior: [
    {
      name: 'reading a file whose worksheet omits sheetFormatPr does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `the read must succeed; got error ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the sheet is discoverable by name after the read',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.length >= 1, 'at least one sheet is exposed');
      },
    },
  ],
} satisfies Case;
