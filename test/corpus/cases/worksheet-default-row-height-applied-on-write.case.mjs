// Cluster: styles
//
// Real-world scenario: a user sets a worksheet's default row height (the height for every row with no
// explicit per-row height) as a worksheet-level property, expecting all otherwise-unstyled rows to
// render at that height. The companion default column width, set the same way, takes effect — but the
// default row height was reported silently dropped, so rows rendered at the application's built-in
// default. The property must be serialized onto the sheet-format definition symmetrically with the
// default column width.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [{name: 'S', properties: {defaultRowHeight: 30, defaultColWidth: 20}, cells: [{ref: 'A1', value: 'x'}]}],
};

export default {
  id: 'worksheet-default-row-height-applied-on-write',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A worksheet default row height set as a property is serialized onto the sheet-format definition ' +
    '(symmetric with the default column width) and round-trips to the same value, rather than being ' +
    'silently dropped so rows fall back to the built-in default.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the default row height is written onto the sheet-format definition',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(sheets.S.sheetFormat.defaultRowHeight, 30, 'the default row height is serialized');
      },
    },
    {
      name: 'the default column width is written (symmetric handling)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(sheets.S.sheetFormat.defaultColWidth, 20, 'the default column width is serialized');
      },
    },
  ],
};
