// Cluster: tables
//
// Real-world scenario: a workbook contains several worksheet tables, each with its own
// autoFilter. A pure read/write round-trip must preserve every table's definition. The table
// references and column counts do survive, but the writer spuriously sets totalsRowShown on
// every table (the source had it off), which — combined with the autoFilter — makes Excel treat
// the tables as corrupt. A round-trip must not turn on a totals row the author never enabled.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'multi-table-autofilter-survives-roundtrip/sample.xlsx';

export default {
  id: 'multi-table-autofilter-survives-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2184},
  cluster: 'tables',
  description:
    'A workbook with multiple tables each carrying an autoFilter round-trips without corruption ' +
    '— every table keeps its autoFilter reference and column count, and the writer does not turn ' +
    'on totalsRowShown that the source left off.',

  behavior: [
    {
      name: 'every table keeps its autoFilter reference and column count (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.roundtripFixtureTableXml(FIXTURE);
        assert.ok(tables.length >= 2, 'the fixture has multiple tables');
        for (const t of tables) {
          assert.ok(t.rewritten, `table ${t.name} still present after round-trip`);
          assert.strictEqual(
            t.rewritten.autoFilterRef,
            t.source.autoFilterRef,
            `autoFilter ref preserved for ${t.name}`,
          );
          assert.strictEqual(
            t.rewritten.columnCount,
            t.source.columnCount,
            `column count preserved for ${t.name}`,
          );
        }
      },
    },
    {
      name: 'no table gains a totalsRowShown the source did not have',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.roundtripFixtureTableXml(FIXTURE);
        for (const t of tables) {
          assert.strictEqual(
            t.rewritten.totalsRowShown,
            t.source.totalsRowShown,
            `totalsRowShown must stay ${t.source.totalsRowShown} for ${t.name}`,
          );
        }
      },
    },
  ],
} satisfies Case;
