// Cluster: sheet-views
//
// Real-world scenario: a worksheet freezes its first (header) row so the header stays visible while
// the rest of the sheet scrolls. This is a frozen sheet view splitting after one row with no column
// split. It serializes as a <pane ySplit="1" ... state="frozen"/> in the sheet view and must survive
// a round-trip so the frozen header is preserved on reopen. Locks the frozen-top-row configuration.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'frozen-top-row-pane-round-trips',
  provenance: {source: 'upstream-issue'},
  cluster: 'sheet-views',
  description:
    'Freezing the first row produces a frozen sheet-view pane split after one row (no column split) ' +
    'and round-trips: after read-back the pane still reports a frozen split of one row and zero columns.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a frozen first-row view emits a pane split after one row',
      baseline: 'pass',
      async expect(api, assert) {
        const {paneEmitted} = await api.frozenTopRowRoundtrip();
        assert.strictEqual(paneEmitted, true, 'the sheet view emits <pane ySplit="1" ... state="frozen"/>');
      },
    },
    {
      name: 'the frozen-top-row split survives a round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {reReadState, reReadYSplit, reReadXSplit} = await api.frozenTopRowRoundtrip();
        assert.strictEqual(reReadState, 'frozen', 'the view is still frozen after read-back');
        assert.strictEqual(reReadYSplit, 1, 'one row stays frozen');
        assert.strictEqual(reReadXSplit, 0, 'no columns are frozen');
      },
    },
  ],
};
