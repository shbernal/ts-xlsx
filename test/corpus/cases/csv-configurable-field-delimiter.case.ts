// Cluster: csv
//
// Real-world scenario: spreadsheet data is exchanged as delimiter-separated text that
// often uses a delimiter other than comma — semicolon is the default in many European
// locales, tab is common for data dumps. Reading such a file, and writing one for such a
// locale, must let the caller set the field delimiter for both directions, and a value
// round-tripped through a non-comma delimiter must keep its field boundaries.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-configurable-field-delimiter',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 108},
  cluster: 'csv',
  description:
    'The CSV field delimiter is configurable on both read and write: a semicolon-delimited ' +
    'file splits on the semicolon when read with that delimiter, a worksheet writes ' +
    'semicolon-separated fields when configured, and a round-trip through a non-comma ' +
    'delimiter preserves field boundaries.',

  behavior: [
    {
      name: 'reading with a configured semicolon delimiter splits fields on the semicolon',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = await api.csvRead({
          csv: 'a;b;c\n1;2;3',
          options: {parserOptions: {delimiter: ';'}},
        });
        assert.ok(ok, 'the read must succeed');
        assert.deepStrictEqual(rows[0], ['a', 'b', 'c'], 'first row split into three fields');
        assert.deepStrictEqual(rows[1], [1, 2, 3], 'numeric fields coerced, split on semicolon');
      },
    },
    {
      name: 'writing with a configured semicolon delimiter emits semicolon-separated fields',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, text} = await api.csvWrite({
          spec: {rows: [['a', 'b', 'c']]},
          options: {formatterOptions: {delimiter: ';'}},
        });
        assert.ok(ok, 'the write must succeed');
        assert.strictEqual(
          text,
          'a;b;c',
          'fields separated by the configured delimiter, not comma',
        );
      },
    },
    {
      name: 'a value round-trips through a non-comma delimiter with field boundaries intact',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {text} = await api.csvWrite({
          spec: {rows: [['x', 'y']]},
          options: {formatterOptions: {delimiter: ';'}},
        });
        const {rows} = await api.csvRead({csv: text, options: {parserOptions: {delimiter: ';'}}});
        assert.deepStrictEqual(
          rows[0],
          ['x', 'y'],
          'the two fields survive the semicolon round-trip',
        );
      },
    },
  ],
} satisfies Case;
