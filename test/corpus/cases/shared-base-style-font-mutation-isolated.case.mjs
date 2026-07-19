// Cluster: styles
//
// Real-world scenario: a user iterates a worksheet's cells, assigns a common base style to each, then
// conditionally overrides one nested property (e.g. font color) on individual cells — green for
// "YES", red for "NO". Instead of only the targeted cells changing, the override bleeds into every
// cell that received the base style, because assigning a shared style object (or shared style id) and
// later mutating a nested property mutates the aliased instance. Correct behavior is copy-on-write:
// assigning a base style then mutating one cell's font must not change its siblings.
//
// This is the authoring-side companion to `loaded-cells-shared-style-object-aliasing` (which covers
// the load-then-mutate path): here two cells are given the SAME base style object, then one cell's
// font is spread-reassigned a color.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'shared-base-style-font-mutation-isolated',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Assigning one base style object to two cells and then mutating one cell’s font color affects ' +
    'only that cell — the sibling given the same base style keeps its original font, with no ' +
    'aliasing bleed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the mutated cell gets the new font color',
      baseline: 'pass',
      async expect(api, assert) {
        const {a1Color} = await api.sharedBaseStyleFontMutation();
        assert.strictEqual(a1Color, 'FF00FF00', 'the targeted cell carries the new color');
      },
    },
    {
      name: 'the sibling sharing the base style does not inherit the mutated font color',
      baseline: 'pass',
      async expect(api, assert) {
        const {a2Color, bled} = await api.sharedBaseStyleFontMutation();
        assert.strictEqual(bled, false, `the sibling must not pick up the color; got A2 color ${JSON.stringify(a2Color)}`);
      },
    },
  ],
};
