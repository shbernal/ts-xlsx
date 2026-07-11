export default {
  id: 'internal-hyperlink-not-written-as-external-relationship',
  cluster: 'styles',
  description:
    'A hyperlink whose destination is inside the same workbook (a "#Sheet2!A1" target) must be ' +
    'serialized as an internal link: a location attribute holding the in-workbook reference and NO ' +
    'external relationship. Writing it as an external-mode r:id relationship (TargetMode="External") ' +
    'AND a location makes a strict consumer resolve both and render the target doubled — the ' +
    'observed "#Sheet2!A1##Sheet2!A1". The internal and external forms must serialize distinctly.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'an internal (#-prefixed) hyperlink carries a location attribute, not an r:id',
      baseline: 'fail',
      async expect(api, assert) {
        const {hyperlinkHasRid, hyperlinkLocation} = await api.internalHyperlinkSerializationReport();
        assert.strictEqual(hyperlinkHasRid, false, 'internal link must not use an r:id relationship');
        assert.ok(hyperlinkLocation, 'internal link must carry a location attribute');
      },
    },
    {
      name: 'an internal hyperlink creates no external-mode worksheet relationship',
      baseline: 'fail',
      async expect(api, assert) {
        const {hasWorksheetRels, relTargetMode} = await api.internalHyperlinkSerializationReport();
        assert.strictEqual(
          hasWorksheetRels && relTargetMode === 'External',
          false,
          'internal link must not emit an External hyperlink relationship',
        );
      },
    },
    {
      name: 'the internal target survives a round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {reReadHyperlink} = await api.internalHyperlinkSerializationReport();
        assert.strictEqual(reReadHyperlink, '#Sheet2!A1');
      },
    },
  ],
};
