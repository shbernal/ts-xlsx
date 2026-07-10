// Cluster: styles
//
// Real-world scenario: a user attaches a border style (a right border) to exactly one column via that
// column's style, and declares later columns with only a width. The border must appear only on the
// declaring column's cells — not bleed into subsequent columns. Correct behavior is that each
// column's declared style is independent; setting a style on one column does not affect any other.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'column-border-style-scoped-to-declaring-column',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A border style declared on one column applies only to that column’s cells; later columns with ' +
    'no style of their own get no border — column styles are independent, not bled into subsequent ' +
    'columns.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the styled column’s cell carries the border',
      baseline: 'pass',
      async expect(api, assert) {
        const {a1} = await api.columnBorderScopedReport();
        assert.strictEqual(a1, true, 'the declaring column’s cell has the right border');
      },
    },
    {
      name: 'columns without a style get no border',
      baseline: 'pass',
      async expect(api, assert) {
        const {b1, c1} = await api.columnBorderScopedReport();
        assert.strictEqual(b1, false, 'the next column does not inherit the border');
        assert.strictEqual(c1, false, 'a further column does not inherit the border');
      },
    },
  ],
};
