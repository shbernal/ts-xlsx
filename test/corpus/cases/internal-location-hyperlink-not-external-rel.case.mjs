// Cluster: address-decoding
//
// Real-world scenario: an author builds a table-of-contents sheet whose cells link to other locations
// inside the same workbook. Each link's hyperlink value begins with '#' — e.g. "#'Target'!A1" —
// marking it as an internal document location rather than an external URL. An internal link must be
// written with a `location` attribute on the <hyperlink> element and NO external relationship; if it
// is instead emitted as an external-mode relationship (r:id → .rels TargetMode="External"), strict
// consumers treat the file as malformed and the link does not navigate. The quoted sheet name must be
// preserved in the location string.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'internal-location-hyperlink-not-external-rel',
  provenance: {source: 'upstream-issue'},
  cluster: 'address-decoding',
  description:
    'A hyperlink whose target begins with "#" is written as an internal link with a location ' +
    'attribute and without a spurious external relationship, so it navigates and the package stays ' +
    'valid; the quoted sheet name is preserved in the location.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the internal hyperlink carries a location attribute with the target reference',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasLocation, location} = await api.internalHyperlinkReport();
        assert.strictEqual(hasLocation, true, 'the <hyperlink> has a location attribute');
        assert.ok(location && location.includes('Target') && location.includes('A1'), `the location keeps the reference; got ${JSON.stringify(location)}`);
      },
    },
    {
      name: 'no external relationship is created for the internal target',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasExternalRel} = await api.internalHyperlinkReport();
        assert.strictEqual(hasExternalRel, false, 'an internal "#" target must not produce an external-mode relationship');
      },
    },
    {
      name: 'the workbook with the internal hyperlink reloads (valid package)',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadOk} = await api.internalHyperlinkReport();
        assert.strictEqual(reloadOk, true, 'the package is valid and reloads');
      },
    },
  ],
};
