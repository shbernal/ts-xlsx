// Cluster: csv
//
// Real-world scenario: a CSV field contains only whitespace (a run of spaces between two delimiters,
// e.g. `firstValue,   ,secondValue`). JavaScript's `Number('   ')` is 0, so a naive type-inference
// mapper turns a whitespace-only field into the number 0 — silently destroying the original content
// and injecting a spurious numeric value. A whitespace-only field is not a number; it must be
// preserved as its literal string. An empty field (nothing between delimiters) stays distinct (null),
// and genuinely numeric fields still parse as numbers.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'csv-whitespace-only-cell-preserved-as-string',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'A CSV field containing only whitespace is preserved as a string rather than coerced to the ' +
    'number 0, while empty fields stay null and genuine numbers still parse numerically.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a whitespace-only field is not coerced to the number 0',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.csvRead({csv: 'firstValue,   ,secondValue\n'});
        assert.notStrictEqual(rows[0][1], 0, 'a whitespace-only field must not become numeric 0');
        assert.strictEqual(typeof rows[0][1], 'string', 'a whitespace-only field is preserved as a string');
      },
    },
    {
      name: 'a genuine numeric field still parses as a number (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.csvRead({csv: 'firstValue,42,secondValue\n'});
        assert.strictEqual(rows[0][1], 42, 'a real number is still coerced to a number');
      },
    },
    {
      name: 'an empty field stays distinct from a whitespace-only field and is not 0 (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.csvRead({csv: 'firstValue,,secondValue\n'});
        assert.notStrictEqual(rows[0][1], 0, 'an empty field must not become numeric 0');
      },
    },
  ],
};
