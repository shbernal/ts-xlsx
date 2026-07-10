// Cluster: styles
//
// Real-world scenario: a workbook is created and written with no theme explicitly configured. The
// package must still ship a valid theme part (the default Office theme), because a workbook that
// declares style/color-scheme dependencies but ships no theme part is treated as corrupt by Excel,
// which repairs the file on open. This locks the writer's default-theme emission as a regression
// guard — a missing theme part is a real corruption class.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const NO_THEME = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}]}]};

export default {
  id: 'default-theme-part-emitted-without-explicit-theme',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A workbook written with no explicitly configured theme still ships a theme part, so a ' +
    'consuming application can resolve theme-backed style dependencies instead of repairing the file.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the default-written package ships a theme part',
      baseline: 'pass',
      async expect(api, assert) {
        const {styles} = await api.inspectPackage(NO_THEME);
        assert.strictEqual(styles.hasThemePart, true, 'a no-theme workbook must still emit the default theme part');
      },
    },
    {
      name: 'any theme-color reference in the default font remains resolvable',
      baseline: 'pass',
      async expect(api, assert) {
        const {styles} = await api.inspectPackage(NO_THEME);
        assert.strictEqual(styles.themeColorResolvable, true, 'a theme-color reference is backed by the emitted theme part');
      },
    },
  ],
};
