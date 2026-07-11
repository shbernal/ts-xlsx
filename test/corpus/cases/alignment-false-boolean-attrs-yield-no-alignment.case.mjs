export default {
  id: 'alignment-false-boolean-attrs-yield-no-alignment',
  cluster: 'styles',
  description:
    'A cell whose alignment element carries only an explicit-false boolean attribute — the ' +
    'wrapText="0" / shrinkToFit="0" that real producers (Excel among them) write out — must read ' +
    'back as having no alignment, not as an { wrapText:false } / { shrinkToFit:false } object. The ' +
    'raw XML attribute value "0" is a truthy JS string, so a reader that guards on the raw value ' +
    'rather than the parsed boolean mistakes an all-false alignment for a present one.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'a cell whose only alignment attribute is wrapText="0" reads back with no alignment',
      baseline: 'fail',
      async expect(api, assert) {
        const {wrapTextZero} = await api.alignmentFalseBooleanReport();
        assert.strictEqual(wrapTextZero, null);
      },
    },
    {
      name: 'a cell whose only alignment attribute is shrinkToFit="0" reads back with no alignment',
      baseline: 'fail',
      async expect(api, assert) {
        const {shrinkZero} = await api.alignmentFalseBooleanReport();
        assert.strictEqual(shrinkZero, null);
      },
    },
  ],
};
