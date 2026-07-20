import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'list-validation-cross-sheet-formula-reference',
  cluster: 'data-validation',
  description:
    'A list data-validation whose source range lives on another sheet (e.g. a dropdown on ' +
    'Sheet1!A1 fed by Sheet2!$D$3:$D$5) must be read back. Spreadsheet applications serialize a ' +
    'cross-worksheet list source in the 2009 `x14` data-validation extension (under the worksheet ' +
    '`extLst`) rather than the plain `<dataValidation>` element; a reader that only understands the ' +
    'standard form silently drops the rule, so the cell reports no validation and a read→write ' +
    'round-trip loses the dropdown entirely. Same-sheet list validations use the standard form and ' +
    'already survive — that contrast is the tell.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'a same-sheet list validation (standard <dataValidation>) is surfaced on its cell',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.readFixtureValidations(
          'list-validation-cross-sheet-formula-reference/cross-sheet-list.xlsx',
        );
        const local = cells['Sheet1!A2'];
        assert.ok(local, 'expected the same-sheet list validation on Sheet1!A2 to be read');
        assert.strictEqual(local.type, 'list');
      },
    },
    {
      name: 'the standard same-sheet validation survives a read→write round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {totalStandard} = await api.roundtripFixtureValidationXml(
          'list-validation-cross-sheet-formula-reference/cross-sheet-list.xlsx',
        );
        assert.ok(
          totalStandard >= 1,
          'expected at least one standard <dataValidation> after round-trip',
        );
      },
    },
    {
      name: 'a cross-sheet list validation (x14 extension) is surfaced on its cell with the foreign-sheet source',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.readFixtureValidations(
          'list-validation-cross-sheet-formula-reference/cross-sheet-list.xlsx',
        );
        const cross = cells['Sheet1!A1'];
        assert.ok(cross, 'expected the cross-sheet list validation on Sheet1!A1 to be read');
        assert.strictEqual(cross.type, 'list');
        const formula = (cross.formulae || []).join(' ');
        assert.match(formula, /Sheet2!/, 'expected the list source to name the foreign sheet');
      },
    },
    {
      name: 'the cross-sheet validation survives a read→write round-trip (x14 extension preserved)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {totalExt} = await api.roundtripFixtureValidationXml(
          'list-validation-cross-sheet-formula-reference/cross-sheet-list.xlsx',
        );
        assert.ok(
          totalExt >= 1,
          'expected the x14 data-validation extension to be re-emitted after round-trip',
        );
      },
    },
  ],
} satisfies Case;
