// Cluster: csv
//
// Real-world scenario: exporting a worksheet to CSV, cells that hold real Date values
// should render using a caller-supplied format (e.g. MM/DD/YYYY) rather than a full ISO-8601
// timestamp with timezone offset. A user exporting dates wants "01/05/2018", not
// "2018-01-05T12:00:00-06:00". A UTC flag controls whether the formatted value reflects
// local time or UTC; without it the local timezone can shift the calendar day.
//
// The date must be a genuine Date value in the cell; the earlier confusion in the wild was
// a cell holding a date *string*, which no format option can reach.
//
// The format is an Excel number-format code -- the same vocabulary `Cell.numFmt` holds, and the only
// one this library speaks for dates. It used to be a moment.js-style token set instead, where the
// month is `MM` and `mm` is minutes, so a caller passing the library's own format code got
// `2024-45-dd` back with nothing reported. The two spellings agree on `mm/dd/yyyy` by coincidence,
// which is exactly why the case below asserts a code where they do not.

import type {Assert, Case, CorpusApi} from '../case.ts';

// A fixed UTC instant so the assertions are timezone-stable regardless of the runner.
const NOON_JAN5_2018_UTC = '2018-01-05T12:00:00.000Z';

export default {
  id: 'csv-write-date-format-honored',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 689},
  cluster: 'csv',
  description:
    'When writing CSV, a Date cell renders using a caller-supplied Excel number-format code, in ' +
    'UTC when requested, instead of always emitting a full ISO-8601 timestamp.',

  behavior: [
    {
      name: 'a Date cell written with a UTC dateFormat renders in that format',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateFormat: 'mm/dd/yyyy', dateUTC: true},
        });
        assert.ok(ok, 'the write must succeed');
        assert.strictEqual(text, '01/05/2018', 'the date renders in the requested format, in UTC');
      },
    },
    {
      name: "the format is a number-format code, so a cell's own numFmt renders what the cell shows",
      async expect(api: CorpusApi, assert: Assert) {
        // The lowercase Excel vocabulary, where a moment.js token set reads `mm` as minutes and the
        // month as `MM`: under that reading this renders `2018-00-dd`, and did.
        const {text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateFormat: 'yyyy-mm-dd', dateUTC: true},
        });
        assert.strictEqual(text, '2018-01-05');
      },
    },
    {
      name: 'minutes and months are told apart by position, as Excel tells them apart',
      async expect(api: CorpusApi, assert: Assert) {
        // One letter, two meanings: `m` is minutes after an hour run or before a seconds run, and a
        // month anywhere else. A token table cannot express that, which is why the old one had to
        // spell the two differently and disagree with every other format code in the library.
        const {text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateFormat: 'yyyy-mm-dd hh:mm:ss', dateUTC: true},
        });
        assert.strictEqual(text, '2018-01-05 12:00:00');
      },
    },
    {
      name: 'a Date cell written with no dateFormat renders as a full ISO-8601 timestamp',
      async expect(api: CorpusApi, assert: Assert) {
        const {text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateUTC: true},
        });
        assert.match(
          text!,
          /^2018-01-05T12:00:00/,
          `default is a full ISO timestamp; got ${JSON.stringify(text)}`,
        );
      },
    },
  ],
} satisfies Case;
