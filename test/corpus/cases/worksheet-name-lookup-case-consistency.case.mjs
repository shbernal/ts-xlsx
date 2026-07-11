export default {
  id: 'worksheet-name-lookup-case-consistency',
  cluster: 'types',
  description:
    'Worksheet-name lookup and worksheet-name uniqueness must agree on what counts as the same ' +
    'name. Adding a worksheet rejects a case-variant duplicate (adding "sheet" when "Sheet" ' +
    'exists throws), but looking one up matches case-sensitively — so a caller who checks ' +
    'getWorksheet("sheet"), sees nothing, and then adds it hits a surprise throw. A name reported ' +
    'absent by lookup must be safely addable; the two APIs cannot disagree on identity.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'the exact-case name is found after the sheet is added',
      baseline: 'pass',
      expect(api, assert) {
        const {foundExact} = api.worksheetNameLookupReport();
        assert.strictEqual(foundExact, true);
      },
    },
    {
      name: 'lookup and add agree: a case-variant name reported absent by getWorksheet is addable',
      baseline: 'fail',
      expect(api, assert) {
        const {foundVariant, addVariantThrew} = api.worksheetNameLookupReport();
        // A consistent API cannot both fail to find "sheet" AND refuse to add it.
        assert.strictEqual(
          foundVariant || !addVariantThrew,
          true,
          'getWorksheet missed the case-variant name yet addWorksheet rejected it as a duplicate',
        );
      },
    },
  ],
};
