// Cluster: types
//
// Real-world scenario: a user has values that look numeric but must be treated as text — an
// identifier, or a formatted amount like "1000.80" where the trailing zero is significant. They set
// the cell value to a JavaScript string. On write-then-read the cell must remain a string carrying
// exactly the characters supplied ("1000.80"), not be coerced into a number that drops the trailing
// zero (yielding 1000.8) or changes its type.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [{name: 'S', cells: [{ref: 'A1', value: '1000.80'}, {ref: 'A2', value: '007'}]}],
};

export default {
  id: 'numeric-looking-string-preserved-as-string',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A cell assigned a numeric-looking JavaScript string ("1000.80", "007") round-trips as a string ' +
    'with its exact characters — including a significant trailing zero and leading zeros — rather ' +
    'than being coerced to a number that changes the value or type.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a string that looks like a decimal keeps its exact characters (trailing zero)',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.cells.A1.value, '1000.80', 'the trailing zero is preserved, not dropped to 1000.8');
      },
    },
    {
      name: 'a string with leading zeros is preserved verbatim',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.cells.A2.value, '007', 'leading zeros survive, not coerced to 7');
      },
    },
    {
      name: 'the preserved value is a string, not a number',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(typeof model.sheets.S.cells.A1.value, 'string', 'the cell reads back as a string type');
      },
    },
  ],
};
