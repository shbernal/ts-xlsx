// Cluster: tables
//
// Real-world scenario: Excel's table-style gallery offers "None" — a table with no banding or theme
// applied. A caller declaring a table with style theme 'None' expects an unstyled table: in OOXML
// that is a tableStyleInfo with no name attribute (exactly what a null/absent theme produces). The
// bug: the writer emits the literal name="None", a reference to a style that does not exist, rather
// than omitting the name. Style flags like showRowStripes set alongside the theme must survive either
// way.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'table-style-none-produces-unstyled-table',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table styled with theme "None" produces an unstyled table — a tableStyleInfo with no name ' +
    'attribute, the same as a null/absent theme — rather than a bogus name="None" that references a ' +
    'non-existent style; style flags set alongside it survive.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a real built-in theme emits its style name (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {real} = await api.tableStyleThemeReport();
        assert.strictEqual(real.ok, true, 'the table writes');
        assert.strictEqual(real.name, 'TableStyleMedium2', 'a real theme emits its name');
      },
    },
    {
      name: 'a null/absent theme emits no style name (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {nullTheme} = await api.tableStyleThemeReport();
        assert.strictEqual(nullTheme.name, null, 'an unstyled table carries no name attribute');
      },
    },
    {
      name: 'the "None" theme produces an unstyled table with no name',
      baseline: 'pass',
      async expect(api, assert) {
        const {none} = await api.tableStyleThemeReport();
        assert.strictEqual(
          none.name,
          null,
          `theme "None" must mean unstyled (no name), not a bogus name="${none.name}" referencing a non-existent style`,
        );
      },
    },
    {
      name: 'style flags set alongside the "None" theme survive',
      baseline: 'pass',
      async expect(api, assert) {
        const {none} = await api.tableStyleThemeReport();
        assert.strictEqual(
          none.hasStripes,
          true,
          'showRowStripes is preserved regardless of the theme value',
        );
      },
    },
  ],
};
