// Cluster: tables
//
// Real-world scenario: an Excel-authored workbook contains a table that has NO autoFilter, and
// whose header-row configuration is the default. A pure read/write round-trip (no logical edits)
// must not alter the table part. Today the writer corrupts it: it injects an autoFilter element
// (with an empty, self-closed filterColumn), flips the header-row setting off, and turns on
// totalsRowShown — producing an internally inconsistent table + autoFilter that Excel treats as
// corrupt and repairs by stripping the Table and AutoFilter features. A round-trip must leave a
// valid table valid.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'table-without-autofilter-not-corrupted-on-roundtrip/sample.xlsx';

export default {
  id: 'table-without-autofilter-not-corrupted-on-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2585},
  cluster: 'tables',
  description:
    'A no-op read/write round-trip of a table that has no autoFilter does not inject an ' +
    'autoFilter, flip the header-row setting off, or turn on totalsRowShown — the re-emitted ' +
    'table part stays valid instead of being repaired away by Excel.',

  behavior: [
    {
      name: 'a table with no autoFilter does not gain one on round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.roundtripFixtureTableXml(FIXTURE);
        const t = tables[0];
        assert.strictEqual(
          t.source.hasAutoFilter,
          false,
          'precondition: the source table has no autoFilter',
        );
        assert.strictEqual(
          t.rewritten.hasAutoFilter,
          false,
          'the round-trip must not inject an autoFilter',
        );
      },
    },
    {
      name: 'the header-row configuration is not flipped on round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.roundtripFixtureTableXml(FIXTURE);
        const t = tables[0];
        assert.strictEqual(
          t.rewritten.headerRowCount,
          t.source.headerRowCount,
          `headerRowCount must be preserved (${t.source.headerRowCount})`,
        );
      },
    },
    {
      name: 'totalsRowShown is not spuriously turned on',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.roundtripFixtureTableXml(FIXTURE);
        const t = tables[0];
        assert.strictEqual(
          t.rewritten.totalsRowShown,
          t.source.totalsRowShown,
          `totalsRowShown must stay ${t.source.totalsRowShown}`,
        );
      },
    },
  ],
} satisfies Case;
