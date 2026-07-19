// Cluster: styles
//
// Real-world scenario: a user types a value into a cell and saves without any explicit formatting.
// On read-back the cell carries no explicit style record, but the workbook defines a default cell
// style (index 0) establishing the default font a viewer would render. Cells (and rows/columns) with
// no explicit style id should resolve to this default so font information is available rather than
// undefined. The reported root cause: the default style at index 0 is loaded but only applied to
// entities with a truthy style id, leaving unstyled cells (and an explicit style id of 0) without a
// resolved font.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'default-font-applies-to-unstyled-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A cell with a plain value and no explicit formatting resolves to the workbook default font (a ' +
    'concrete name/size) on read-back, rather than reporting no font at all.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an unstyled cell resolves to a default font (not undefined)',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasFont, fontName} = await api.unstyledCellFontReport();
        assert.strictEqual(hasFont, true, 'the unstyled cell must resolve to the workbook default font');
        assert.ok(fontName, `the default font has a concrete name; got ${JSON.stringify(fontName)}`);
      },
    },
    {
      name: 'the resolved default font carries a concrete size',
      baseline: 'pass',
      async expect(api, assert) {
        const {fontSize} = await api.unstyledCellFontReport();
        assert.ok(typeof fontSize === 'number' && fontSize > 0, `the default font has a numeric size; got ${JSON.stringify(fontSize)}`);
      },
    },
  ],
};
