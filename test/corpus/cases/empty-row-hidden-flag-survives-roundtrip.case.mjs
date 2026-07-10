// Cluster: xlsx-io
//
// Real-world scenario: a user sets a row's hidden property to true but never assigns any cell value
// in that row (an intentional blank spacer/hidden row). On write-then-read the hidden state is lost —
// the row comes back visible — because the writer only emits a <row> element for rows that carry
// cells or another materialized property. A blank row that carries a row-level property (hidden,
// height, outline level) must still be written so that property survives.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'empty-row-hidden-flag-survives-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'A blank row (no cell values) that is set hidden stays hidden after a round-trip; a blank row ' +
    'given a height or outline level likewise retains it — the row-level property is written even ' +
    'without cell content.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a blank row set hidden stays hidden after a round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {row3Hidden} = await api.hiddenEmptyRowReport();
        assert.strictEqual(row3Hidden, true, 'the hidden flag on a content-less row must survive the write');
      },
    },
    {
      name: 'a blank row with a height keeps both its hidden flag and height (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {row4Hidden, row4Height} = await api.hiddenEmptyRowReport();
        assert.strictEqual(row4Hidden, true, 'a blank hidden row with a height stays hidden');
        assert.strictEqual(row4Height, 25, 'the height survives');
      },
    },
    {
      name: 'a blank row with an outline level stays hidden after a round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {row5Hidden} = await api.hiddenEmptyRowReport();
        assert.strictEqual(row5Hidden, true, 'a blank hidden row with an outline level stays hidden');
      },
    },
  ],
};
