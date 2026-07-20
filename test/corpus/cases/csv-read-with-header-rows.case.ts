// Cluster: csv
//
// Real-world scenario: reading a CSV with the parser's header mode enabled, the parser
// emits each data row as an object keyed by header name rather than as an array. A reader
// that unconditionally runs a per-cell map over each row assumes an array and throws
// "data.map is not a function" the moment header mode is on — so the entire header-mode
// read path is broken. Reading the same CSV without header mode (every line an array of
// fields) works and is the control.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-read-with-header-rows',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1695},
  cluster: 'csv',
  description:
    'Reading a CSV with the parser’s header mode enabled must not throw: header-keyed row ' +
    'objects must be handled rather than assumed to be arrays. Reading without header mode ' +
    'splits every line into array fields and is the working control.',

  behavior: [
    {
      name: 'reading with header mode enabled does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.csvRead({
          csv: 'name,age\nalice,30',
          options: {parserOptions: {headers: true}},
        });
        assert.ok(ok, `header-mode read must not throw; got error ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'in header mode the header row is consumed and data rows carry the values',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = await api.csvRead({
          csv: 'name,age\nalice,30',
          options: {parserOptions: {headers: true}},
        });
        assert.ok(ok, 'header-mode read must complete to yield rows');
        assert.strictEqual(rows.length, 1, 'the header line is consumed, leaving one data row');
        assert.deepStrictEqual(
          rows[0],
          ['alice', 30],
          'the data row carries the values in column order',
        );
      },
    },
    {
      name: 'reading without header mode splits every line into array fields (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = await api.csvRead({csv: 'name,age\nalice,30', options: {}});
        assert.ok(ok, 'array-mode read works');
        assert.deepStrictEqual(
          rows[0],
          ['name', 'age'],
          'the header line is itself a data row in array mode',
        );
        assert.deepStrictEqual(rows[1], ['alice', 30], 'the second line splits into fields');
      },
    },
  ],
} satisfies Case;
