// Cluster: csv
//
// Real-world scenario: reading a CSV, fields that merely *resemble* a date must not be
// silently turned into Date values. Identifiers and codes like "2020-00001" (a padded ID),
// "1-3" (an inventory code), or "3-4" (a range label) are text and must stay text —
// coercing them corrupts the data and can crash downstream code. Genuinely numeric fields
// should still become numbers, and genuine ISO dates should still become dates. The reader
// must coerce conservatively: only clear numbers and strictly-formatted dates convert.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-read-preserves-identifier-strings',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 711},
  cluster: 'csv',
  description:
    'CSV read coerces conservatively: an identifier or code that resembles a date ' +
    '("2020-00001", "1-3") is preserved as a string, while a clearly numeric field becomes ' +
    'a number and a strictly-formatted ISO date becomes a Date.',

  behavior: [
    {
      name: 'identifier-like strings that resemble dates are preserved as text',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = await api.csvRead({csv: '2020-00001,1-3,3-4', options: {}});
        assert.ok(ok, 'the read must succeed');
        assert.deepStrictEqual(
          rows[0],
          ['2020-00001', '1-3', '3-4'],
          'padded IDs and dash-codes stay strings, not coerced to dates',
        );
      },
    },
    {
      name: 'a clearly numeric field is read as a number',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '123,45.6', options: {}});
        assert.deepStrictEqual(rows[0], [123, 45.6], 'numeric fields coerce to numbers');
      },
    },
    {
      name: 'a strictly-formatted ISO date field is read as a Date',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '2018-01-05', options: {}});
        const cell = rows[0][0];
        assert.ok(
          cell && typeof cell === 'object' && cell.date,
          `a real date should coerce; got ${JSON.stringify(cell)}`,
        );
        assert.match(
          cell.date,
          /^2018-01-0[45]T/,
          'the parsed date is Jan 5 2018 (modulo timezone)',
        );
      },
    },
  ],
} satisfies Case;
