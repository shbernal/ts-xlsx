// Cluster: styles
//
// Real-world scenario: the default font a writer emits references its color by
// *theme index* (e.g. <color theme="1"/>) rather than a literal ARGB. That is valid
// OOXML only if the package also ships a theme part that defines the color slots —
// otherwise Excel cannot resolve the color and prompts to repair the file on open.
// The invariant: whenever a serialized font references a theme color, a theme part
// backing it is present in the package.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const SPEC = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}]}]};

export default {
  id: 'theme-color-font-backed-by-theme-part',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 425},
  cluster: 'styles',
  description:
    'A font that references a color by theme index is backed by a theme part in the ' +
    'written package, so Excel can resolve the color instead of repairing the file.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the written package ships a theme part',
      baseline: 'pass',
      async expect(api, assert) {
        const {styles} = await api.inspectPackage(SPEC);
        assert.strictEqual(styles.hasThemePart, true, 'package should include a theme part');
      },
    },
    {
      name: 'any theme-color reference in the default font is resolvable',
      baseline: 'pass',
      async expect(api, assert) {
        const {styles} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          styles.themeColorResolvable,
          true,
          'a theme-color font reference must be backed by a theme part'
        );
      },
    },
  ],
};
