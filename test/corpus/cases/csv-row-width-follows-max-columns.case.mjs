// Cluster: csv
//
// Real-world scenario: a sheet's first row is narrower than a later row (e.g. a one-cell title row
// above three-column data). When exporting to CSV, every populated cell must survive: the number of
// comma-separated fields per line follows the sheet's maximum column extent across all rows, not the
// width of the first row. A writer that sized every row to the first row's width would silently drop
// the extra fields of wider rows. This locks the correct behavior as a regression guard.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const NARROW_FIRST = {rows: [['a'], ['b', 'c', 'd']]};

export default {
  id: 'csv-row-width-follows-max-columns',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'CSV export sizes each row by the sheet’s maximum column extent, so a wide row following a ' +
    'narrow first row keeps all its fields instead of being truncated to the first row’s width.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a three-cell row after a one-cell row emits three fields',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, text} = await api.csvWrite({spec: NARROW_FIRST});
        assert.strictEqual(ok, true, 'the CSV writes');
        const lines = String(text).split(/\r?\n/).filter(Boolean);
        assert.strictEqual(
          lines[1],
          'b,c,d',
          'the wider row keeps all three fields, not truncated to one',
        );
      },
    },
    {
      name: 'no populated cell is dropped because a preceding row was narrower',
      baseline: 'pass',
      async expect(api, assert) {
        const {text} = await api.csvWrite({spec: NARROW_FIRST});
        const lines = String(text).split(/\r?\n/).filter(Boolean);
        assert.strictEqual(
          lines[1].split(',').length,
          3,
          'the per-row field count reflects the max column extent',
        );
      },
    },
  ],
};
