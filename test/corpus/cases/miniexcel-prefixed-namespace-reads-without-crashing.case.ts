// Cluster: xlsx-io
//
// Real-world scenario: a workbook produced by a foreign generator (MiniExcel) serializes every
// SpreadsheetML element under an explicit namespace prefix (<x:workbook>, <x:sheetData>,
// <x:c>…), uses inline `t="str"` cell values, and points its worksheet relationship at an
// absolute package path. The file is spec-valid and opens in desktop apps, but the library
// crashes reading it ("Cannot set properties of undefined (setting 'sheetNo')") — it assumes the
// unprefixed element names. A reader must tolerate the prefixed serialization: read the file
// without throwing, expose the single sheet by its declared name, and read its inline cell text.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'miniexcel-prefixed-namespace-reads/sample.xlsx';

export default {
  id: 'miniexcel-prefixed-namespace-reads-without-crashing',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2544},
  cluster: 'xlsx-io',
  description:
    'A workbook whose SpreadsheetML is serialized under an explicit namespace prefix (x:workbook, ' +
    'x:sheetData, x:c) with inline string cells reads without crashing and exposes its declared ' +
    'sheet, instead of throwing on an assumed unprefixed element name.',

  behavior: [
    {
      name: 'the prefixed-namespace workbook reads without throwing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `the read must not crash; got error ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the declared sheet name is exposed',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(
          sheetNames,
          ['apis 3445'],
          'the single sheet is discovered by its declared name',
        );
      },
    },
  ],
} satisfies Case;
