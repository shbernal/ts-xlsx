// Cluster: styles
//
// Real-world scenario: a caller sets a cell's number format from a structured object (an object
// carrying an id and a format-code field, e.g. copied from another cell's parsed numFmt) instead of a
// plain format-code string. The writer must not blindly stringify that object into the styles part's
// formatCode attribute — doing so emits formatCode="[object Object]", a malformed number format that
// Excel reports as a corrupt package on open. A legitimately-set format-code STRING, meanwhile, must
// survive a write even when the cell also carries alignment, font, and protection styling.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'numfmt-object-value-corrupts-styles',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A non-string number format assigned to a cell must not serialize to formatCode="[object Object]" ' +
    'in the styles part (which corrupts the package), while a valid format-code string survives a ' +
    'write even alongside alignment, font, and protection styling.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a valid format-code string survives a write alongside other style facets',
      baseline: 'pass',
      async expect(api, assert) {
        const {controlNumFmtReload} = await api.numFmtObjectCorruptionReport();
        assert.strictEqual(
          controlNumFmtReload,
          'yyyy-mmm-dd',
          `the string number format must survive the write; got ${JSON.stringify(controlNumFmtReload)}`
        );
      },
    },
    {
      name: 'an object-valued number format is not stringified into the styles part as "[object Object]"',
      baseline: 'fail',
      async expect(api, assert) {
        const {stylesHasObjectObject} = await api.numFmtObjectCorruptionReport();
        assert.strictEqual(
          stylesHasObjectObject,
          false,
          'the styles part must never contain formatCode="[object Object]" from a coerced object numFmt'
        );
      },
    },
  ],
};
