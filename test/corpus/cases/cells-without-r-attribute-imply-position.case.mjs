// Cluster: address-decoding
//
// Real-world scenario: the `r` position attribute on <row> and <c> elements is optional in
// SpreadsheetML — a generator may omit it, in which case position is implied by document order
// (the first row is row 1, and cells fill columns A, B, C… left to right). A reader that requires
// `r` fails with "Invalid row number in model". A worksheet whose rows and cells omit `r` must
// read successfully, with each element's position inferred from its order.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'cells-without-r-attribute-imply-position/sample.xlsx';

export default {
  id: 'cells-without-r-attribute-imply-position',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2961},
  cluster: 'address-decoding',
  description:
    'A worksheet whose <row>/<c> elements omit the optional r position attribute reads without ' +
    'error, with row and column positions implied by document order.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a worksheet with cells lacking the r attribute reads without error',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `implied positions must not be a fatal error; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the worksheet is accessible after inferring positions',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.length >= 1, 'the sheet is exposed once positions are implied');
      },
    },
  ],
};
