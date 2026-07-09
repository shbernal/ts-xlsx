// Cluster: tables
//
// Real-world scenario: an author attaches a dropdown (a data validation of type "list") to cells.
// There are two supported ways to express the allowed values, and both are common. The first is an
// inline, comma-separated literal wrapped in double quotes — formula `"Male,Female"` — embedding the
// whole choice set directly in the validation. The second is a formula that references a range
// holding the choices — `Levels!$A$2:$A$9999` — letting the dropdown draw from data written on a
// (often hidden) helper sheet, so it behaves "dynamically" as that data changes. Whichever form the
// author supplies must survive a write→read round-trip with its type, target, and formula text
// intact, and the emitted XML must be one a strict consumer (Excel) accepts without a repair prompt.
//
// Durable constraint worth recording: the inline literal form has an Excel length ceiling (~255
// chars); larger option sets silently fail to display in Excel and must use the range-reference
// form. That is a design guardrail for authors, captured in the spec notes — the corpus locks the
// mechanical guarantee that neither form is mangled or coerced on the way through.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'list-validation-value-source-forms-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'List-type data validations round-trip both value-source forms an author uses: an inline quoted ' +
    'literal ("Male,Female") and a cross-sheet range reference (Levels!$A$2:$A$9999), each preserved ' +
    'verbatim, emitting well-formed OOXML with one <dataValidation> per target range.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an inline quoted comma-separated list formula round-trips as a list validation with its literal preserved',
      baseline: 'pass',
      async expect(api, assert) {
        const {readBack, xml} = await api.authorListValidations([
          {ref: 'B2', formula: '"Male,Female"', error: 'pick one'},
        ]);
        assert.strictEqual(readBack.B2.type, 'list', 'the re-read validation must still be a list');
        assert.deepStrictEqual(readBack.B2.formulae, ['"Male,Female"'], 'the inline literal must survive verbatim');
        assert.ok(xml.formula1.includes('"Male,Female"'), `serialized formula1 must carry the literal; got ${JSON.stringify(xml.formula1)}`);
      },
    },
    {
      name: 'a cross-sheet range-reference list formula round-trips with the reference string intact',
      baseline: 'pass',
      async expect(api, assert) {
        const {readBack, xml} = await api.authorListValidations([
          {ref: 'C2', formula: 'Levels!$A$2:$A$9999'},
        ]);
        assert.strictEqual(readBack.C2.type, 'list');
        assert.deepStrictEqual(readBack.C2.formulae, ['Levels!$A$2:$A$9999'], 'the range reference must survive verbatim');
        assert.ok(xml.formula1.includes('Levels!$A$2:$A$9999'), `serialized formula1 must carry the reference; got ${JSON.stringify(xml.formula1)}`);
      },
    },
    {
      name: 'multiple list validations on distinct ranges each serialize as their own well-formed dataValidation',
      baseline: 'pass',
      async expect(api, assert) {
        const {readBack, xml} = await api.authorListValidations([
          {ref: 'B2', formula: '"Male,Female"'},
          {ref: 'C2', formula: 'Levels!$A$2:$A$9999'},
        ]);
        assert.strictEqual(readBack.B2.type, 'list');
        assert.strictEqual(readBack.C2.type, 'list');
        assert.strictEqual(xml.count, 2, 'each distinct target range must emit its own dataValidation');
        assert.ok(xml.wellFormed, 'the emitted dataValidations block must be well-formed OOXML');
      },
    },
  ],
};
