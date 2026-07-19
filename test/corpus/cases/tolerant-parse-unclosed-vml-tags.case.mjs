// Cluster: xlsx-io
//
// Real-world scenario: a workbook's legacy VML drawing part (used for comments/controls) contains
// unclosed break-style tags — HTML-ish <br> without a closing form — as some foreign generators
// emit. A strict XML parser aborts the entire load with an "unexpected close tag" error, so the
// whole workbook is unreadable because of a malformed auxiliary part. A malformed legacy
// VML/drawing part must not abort parsing of the rest of the package: the worksheet and cell data
// must remain readable.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'tolerant-parse-unclosed-vml-tags/sample.xlsx';

export default {
  id: 'tolerant-parse-unclosed-vml-tags',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1958},
  cluster: 'xlsx-io',
  description:
    'A workbook whose legacy VML drawing part contains unclosed break-style tags loads without ' +
    'throwing an "unexpected close tag" error — a malformed auxiliary part does not abort parsing ' +
    'of the worksheet and cell data.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook with malformed VML break tags reads without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(
          ok,
          `the read must not abort on a malformed VML part; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheets remain accessible after tolerating the malformed part',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.length >= 1, 'at least one worksheet is exposed');
      },
    },
  ],
};
