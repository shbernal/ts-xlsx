// Cluster: sheet-views
//
// Real-world scenario: a workbook has a worksheet with frozen panes, and a consumer wants to
// unfreeze it — the natural edit is to replace the sheet's view with a plain normal view. The
// emitted sheetViews XML must then be self-consistent: a normal view carries NO <pane> element and
// no frozen-only pane attributes. If pane markup leaks onto a normal view, a spreadsheet application
// flags the file as needing repair on open. Conversely a genuinely frozen view must still emit its
// <pane>, and the round-trip of an unfrozen sheet must report state 'normal' with no split. This
// pins the view-state edit as producing valid, stable output in each direction.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'unfreeze-frozen-view-produces-valid-normal-view',
  provenance: {source: 'upstream-issue'},
  cluster: 'sheet-views',
  description:
    'Unfreezing a worksheet by replacing its frozen view with a normal view emits valid sheetViews ' +
    'XML — no leftover <pane> element on the normal view — and the reloaded sheet reports state ' +
    "'normal' with no split, while a frozen view still emits its <pane>.",

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a frozen view emits a <pane> element',
      baseline: 'pass',
      async expect(api, assert) {
        const {frozenHasPane} = await api.unfreezeViewRoundtrip();
        assert.strictEqual(frozenHasPane, true, 'a frozen view must serialize a <pane>');
      },
    },
    {
      name: 'unfreezing to a normal view leaves no <pane> element behind',
      baseline: 'pass',
      async expect(api, assert) {
        const {normalHasPane} = await api.unfreezeViewRoundtrip();
        assert.strictEqual(normalHasPane, false, 'a normal view must not carry a leftover <pane> (that markup triggers a repair prompt)');
      },
    },
    {
      name: 'the unfrozen sheet reloads as a normal view with no split',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadedState, reloadedHasSplit} = await api.unfreezeViewRoundtrip();
        assert.strictEqual(reloadedState, 'normal', `unfrozen view must reload as 'normal'; got ${JSON.stringify(reloadedState)}`);
        assert.strictEqual(reloadedHasSplit, false, 'unfrozen view must not retain a frozen split');
      },
    },
  ],
};
