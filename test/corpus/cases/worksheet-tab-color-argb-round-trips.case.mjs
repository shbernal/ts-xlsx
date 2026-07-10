// Cluster: styles
//
// Real-world scenario: a worksheet's tab is colored by supplying an 8-digit ARGB hex string. Colors
// in the format are ARGB — the first two hex digits are the alpha channel — so FFFF0000 is opaque
// red, not a blue-ish RGBA misread. The tab color must round-trip verbatim, and a worksheet with no
// tab color set must not acquire a spurious one. Locks the tab-color ARGB contract.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'worksheet-tab-color-argb-round-trips',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A worksheet tab color given as an 8-digit ARGB (alpha first) round-trips verbatim through a ' +
    'write/read cycle, and a worksheet with no tab color set does not gain one.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the tab color is written as the exact 8-digit ARGB',
      baseline: 'pass',
      async expect(api, assert) {
        const {tabColorArgbWritten} = await api.tabColorRoundtrip();
        assert.strictEqual(tabColorArgbWritten, 'FFFF0000', 'the ARGB is written verbatim (alpha FF first, then red)');
      },
    },
    {
      name: 'the tab color reads back as the same ARGB',
      baseline: 'pass',
      async expect(api, assert) {
        const {reReadArgb} = await api.tabColorRoundtrip();
        assert.strictEqual(reReadArgb, 'FFFF0000', 'the tab color survives the round-trip');
      },
    },
    {
      name: 'a worksheet with no tab color does not acquire one',
      baseline: 'pass',
      async expect(api, assert) {
        const {uncoloredHasTab} = await api.tabColorRoundtrip();
        assert.strictEqual(uncoloredHasTab, false, 'no spurious tab color appears on an uncolored sheet');
      },
    },
  ],
};
