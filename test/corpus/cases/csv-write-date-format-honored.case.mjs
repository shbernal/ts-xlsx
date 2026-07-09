// Cluster: csv
//
// Real-world scenario: exporting a worksheet to CSV, cells that hold real Date values
// should render using a caller-supplied format (e.g. MM/DD/YYYY) rather than a full ISO-8601
// timestamp with timezone offset. A user exporting dates wants "01/05/2018", not
// "2018-01-05T12:00:00-06:00". A UTC flag controls whether the formatted value reflects
// local time or UTC — without it the local timezone can shift the calendar day.
//
// The date must be a genuine Date value in the cell; the earlier confusion in the wild was
// a cell holding a date *string*, which no format option can reach.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// A fixed UTC instant so the assertions are timezone-stable regardless of the runner.
const NOON_JAN5_2018_UTC = '2018-01-05T12:00:00.000Z';

export default {
  id: 'csv-write-date-format-honored',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 689},
  cluster: 'csv',
  description:
    'When writing CSV, a Date cell renders using a caller-supplied date format string, in ' +
    'UTC when requested, instead of always emitting a full ISO-8601 timestamp.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a Date cell written with a UTC dateFormat renders in that format',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateFormat: 'MM/DD/YYYY', dateUTC: true},
        });
        assert.ok(ok, 'the write must succeed');
        assert.strictEqual(text, '01/05/2018', 'the date renders in the requested format, in UTC');
      },
    },
    {
      name: 'a Date cell written with no dateFormat renders as a full ISO-8601 timestamp',
      baseline: 'pass',
      async expect(api, assert) {
        const {text} = await api.csvWrite({
          spec: {rows: [[{date: NOON_JAN5_2018_UTC}]]},
          options: {dateUTC: true},
        });
        assert.match(
          text,
          /^2018-01-05T12:00:00/,
          `default is a full ISO timestamp; got ${JSON.stringify(text)}`
        );
      },
    },
  ],
};
