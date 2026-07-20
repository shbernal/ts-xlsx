// Cluster: streaming
//
// Real-world scenario: a large file is read row-by-row with the streaming reader, and a column
// holds dates (numeric cells whose style resolves to a date number format). The full (buffered)
// read surfaces those cells as Date values; the streaming read surfaces the raw serial number
// instead, because it does not apply cell styles when deciding a cell's type. A date read as a
// bare number breaks every downstream consumer. Streaming read must resolve a date-formatted
// numeric cell to a date, exactly as the full read does — while genuinely numeric and string
// cells keep their types.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'streaming-read-applies-date-format/sample.xlsx';
// A second fixture whose date cells use East-Asian built-in number-format ids whose default entry
// is a *locale-keyed map* of format strings (not a single universal code). The streaming reader's
// default-format lookup reads only the single-code shape, so these built-in date ids resolve to
// nothing and degrade to raw serials — the same date-vs-number failure, via a distinct root cause.
const LOCALE_FIXTURE = 'streaming-read-applies-date-format/locale-dates.xlsx';

export default {
  id: 'streaming-read-applies-date-format',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1430},
  cluster: 'streaming',
  description:
    'The streaming reader applies cell number formats when typing cells, so a date-formatted ' +
    'numeric cell is surfaced as a date (matching the full read) rather than as a raw serial ' +
    'number — while plain numeric and string cells keep their types.',

  behavior: [
    {
      name: 'plain numeric and string cells keep their types in streaming read (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.streamReadFixture(FIXTURE, ['A1', 'C2']);
        assert.strictEqual(cells.A1.type, 'string', 'the header cell is a string');
        assert.strictEqual(cells.C2.type, 'number', 'the numeric cost cell is a number');
      },
    },
    {
      name: 'a date-formatted cell is surfaced as a date in streaming read',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.streamReadFixture(FIXTURE, ['B2']);
        assert.strictEqual(
          cells.B2.type,
          'date',
          `a date-formatted cell must stream as a date, not the raw serial; got ${JSON.stringify(cells.B2)}`,
        );
      },
    },
    {
      name: 'the full read agrees a locale-keyed built-in date id is a date (oracle)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(LOCALE_FIXTURE, ['A2', 'A5']);
        assert.strictEqual(
          cells.A2.type,
          'date',
          'the plain built-in date cell reads as a date in the full read',
        );
        assert.strictEqual(
          cells.A5.type,
          'date',
          'the locale-keyed built-in date cell reads as a date in the full read',
        );
      },
    },
    {
      name: 'a cell using a locale-keyed built-in date format streams as a date, matching the full read',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.streamReadFixture(LOCALE_FIXTURE, ['A5']);
        assert.strictEqual(
          cells.A5.type,
          'date',
          `a built-in locale-keyed date id must stream as a date, not a raw serial; got ${JSON.stringify(cells.A5)}`,
        );
      },
    },
  ],
} satisfies Case;
