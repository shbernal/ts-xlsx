// Cluster: csv
//
// Real-world scenario: CSV has no single row terminator. A file written on a Unix host ends rows with
// LF, one written on Windows with CRLF, and one exported by a classic-Mac-era tool (or by software
// still carrying that convention, such as older Excel-for-Mac exports and some lab instruments) with a
// bare CR. Files that mix them exist too, because they are assembled by concatenation.
//
// A reader that treats a bare CR as a character to discard rather than as a row terminator does not
// report an error: it splices the next row's first field onto this row's last one and loses every row
// boundary in the file. That is silent corruption of both the row count and the field values, and it
// is indistinguishable downstream from a file that really did have one long row.

import type {Assert, Case, CorpusApi} from '../case.ts';

// One row terminator of each kind in a single file, so no case can pass by handling only the
// terminator that happens to appear first.
const MIXED = 'a,b\nc,d\r\ne,f\rg,h';

export default {
  id: 'csv-row-endings-lf-crlf-and-bare-cr',
  cluster: 'csv',
  description:
    'LF, CRLF and a bare CR each terminate a CSV row, so a file mixing all three reads as the rows ' +
    'it states rather than collapsing the CR-separated ones into their neighbours; a CR that is ' +
    'part of a CRLF pair does not terminate a second, empty row.',
  provenance: {source: 'audit'},

  behavior: [
    {
      name: 'a file mixing LF, CRLF and bare-CR endings reads as four rows',
      expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = api.csvRead({csv: MIXED});
        assert.strictEqual(ok, true);
        assert.deepStrictEqual(rows, [
          ['a', 'b'],
          ['c', 'd'],
          ['e', 'f'],
          ['g', 'h'],
        ]);
      },
    },
    {
      name: 'a bare CR terminates its row instead of being spliced into a field',
      expect(api: CorpusApi, assert: Assert) {
        const {rows} = api.csvRead({csv: 'e,f\rg,h'});
        assert.deepStrictEqual(
          rows,
          [
            ['e', 'f'],
            ['g', 'h'],
          ],
          'dropping the CR would concatenate `f` and `g` into one field and lose a row',
        );
      },
    },
    {
      name: 'the CR of a CRLF pair does not terminate a row of its own',
      expect(api: CorpusApi, assert: Assert) {
        const {rows} = api.csvRead({csv: 'a,b\r\nc,d'});
        assert.deepStrictEqual(rows, [
          ['a', 'b'],
          ['c', 'd'],
        ]);
      },
    },
    {
      name: 'a trailing bare CR does not yield a spurious empty row',
      expect(api: CorpusApi, assert: Assert) {
        const {rows} = api.csvRead({csv: 'a,b\r'});
        assert.deepStrictEqual(rows, [['a', 'b']]);
      },
    },
  ],
} satisfies Case;
