// Cluster: tables
//
// Real-world scenario: some generators emit a table whose autoFilter carries a filterColumn whose
// colId points outside the table's declared columns — a dangling column reference. When such a file
// is opened, the reader resolves that colId against the table's column model, finds nothing, and
// crashes ("Cannot set properties of undefined") while trying to mark the filter button, aborting
// the entire workbook load. The reader must tolerate the out-of-range reference — ignore or clamp it
// — and still surface the worksheet and its data, rather than letting one bad index reject a whole
// otherwise-valid package.
//
// The fixture is authored by building a normal table and injecting a filterColumn with an
// out-of-range colId into its table-part autoFilter, reproducing the dangling reference without
// depending on a specific generator to write it.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'table-filter-column-out-of-range/sample.xlsx';

export default {
  id: 'table-filter-column-out-of-range-tolerated',
  provenance: {source: 'upstream-pr'},
  cluster: 'tables',
  description:
    'Loading a workbook whose table autoFilter references a filterColumn colId outside the declared ' +
    'columns completes without throwing and the worksheet survives, rather than the dangling column ' +
    'reference crashing the entire read.',

  behavior: [
    {
      name: 'an out-of-range filterColumn colId does not abort the load',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a dangling filter-column reference must be tolerated, not crash the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheet is recovered intact despite the dangling reference',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(sheetNames, ['Sheet1'], 'the sheet survives the tolerant read');
      },
    },
  ],
} satisfies Case;
