// Cluster: csv
//
// Real-world scenario: a caller picks a delimiter -- a semicolon for a European locale, a tab for a
// paste into a spreadsheet, and occasionally something the format cannot carry, like `||` from a
// pipe-delimited convention or an empty string from an unset config value. The reader refused a
// multi-character delimiter, because its character-scan parser cannot honour one. The writer checked
// nothing at all, in a module whose header is entirely about lossless round-tripping.
//
// So `readCsv(writeCsv(wb, {delimiter: '||'}))` threw on text this same codec had just produced. And
// the empty string was worse than a throw: `quoteField` asks `field.includes(delimiter)`, which is
// true of every field, so every field was quoted and the file went out with no separators in it --
// a file that parses cleanly as a single column and loses the shape of the data with nothing
// reported anywhere.
//
// The rule this locks: one validator, called from both entry points. Whatever the writer accepts,
// the reader reads back.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'the-two-halves-of-the-csv-codec-agree-on-a-delimiter',
  provenance: {source: 'codec-symmetry-audit'},
  cluster: 'csv',
  description:
    'A delimiter the CSV reader cannot honour is refused by the writer as well, with the same ' +
    'RangeError, so every delimiter the writer accepts round-trips through the reader.',

  behavior: [
    {
      name: 'every delimiter the writer accepts reads back as the rows that were written',
      async expect(api: CorpusApi, assert: Assert) {
        const written = (await api.csvDelimiterAgreement()).filter((row) => row.wroteOk);
        assert.equal(written.length, 3, 'comma, semicolon and tab');
        for (const row of written) {
          assert.equal(row.readOk, true, `${row.delimiter}: reads back (${row.readError})`);
          assert.deepEqual(
            row.rows,
            [
              ['a', 'b'],
              ['c', 'd'],
            ],
            `${row.delimiter}: the grid survives the round trip`,
          );
        }
      },
    },
    {
      name: 'a delimiter the reader cannot honour is refused by the writer, not silently emitted',
      async expect(api: CorpusApi, assert: Assert) {
        const rows = new Map(
          (await api.csvDelimiterAgreement()).map((row) => [row.delimiter, row]),
        );
        for (const delimiter of ['||', '']) {
          const row = rows.get(delimiter);
          assert.equal(row?.wroteOk, false, `${JSON.stringify(delimiter)}: the write is refused`);
          assert.match(
            String(row?.writeError),
            /single character/,
            'and says what a delimiter may be',
          );
        }
      },
    },
    {
      name: 'the delimiter actually separates: it is not swallowed by quoting',
      async expect(api: CorpusApi, assert: Assert) {
        const rows = new Map(
          (await api.csvDelimiterAgreement()).map((row) => [row.delimiter, row]),
        );
        assert.equal(rows.get(';')?.text, 'a;b\nc;d');
        assert.equal(rows.get('\t')?.text, 'a\tb\nc\td');
      },
    },
  ],
} satisfies Case;
