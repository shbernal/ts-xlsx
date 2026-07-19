// Cluster: csv
//
// Real-world scenario: a CSV file holds cells that look numeric. Coercing every numeric-looking cell
// through JavaScript's Number() silently corrupts any integer beyond the safe-integer range (and any
// decimal whose digits exceed double precision): a 20-digit id like "56343416020533614003" becomes
// 56343416020533620000, losing digits with no error and no way to recover the original. A reader
// should only produce a number when the numeric conversion round-trips back to the original text;
// otherwise it should keep the cell as its original string so no data is lost. Ordinary in-range
// numbers must still parse as numbers.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const BIG = '56343416020533614003';
const CSV = `id,val\n${BIG},42\n1.5,7\n`;

export default {
  id: 'csv-large-number-precision-preserved-as-string',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'A CSV numeric string that exceeds double precision is preserved as its original string (every ' +
    'digit retained) rather than coerced through Number() to a rounded value, while an ordinary ' +
    'in-range number still parses as a number.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an ordinary in-range number is still parsed as a number',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.csvRead({csv: CSV});
        assert.strictEqual(rows[1][1], 42, 'a small integer parses as a number');
        assert.strictEqual(rows[2][0], 1.5, 'a decimal parses as a number');
      },
    },
    {
      name: 'an over-precision numeric string is preserved with all its digits',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.csvRead({csv: CSV});
        assert.strictEqual(
          rows[1][0],
          BIG,
          `a 20-digit value must be preserved verbatim, not rounded through Number(); got ${JSON.stringify(rows[1][0])}`
        );
      },
    },
  ],
};
