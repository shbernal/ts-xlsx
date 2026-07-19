// Cluster: core-model
//
// Real-world scenario: a program opens a hand-crafted *template* `.xlsx` that already
// contains data validations (typically dropdown lists), fills in some cells, and writes
// the result back out. The written file must still carry the template's validations — a
// dropdown the template author created must keep working in the output.
//
// The trap: modern Excel stores many validations — list validations that reference
// another sheet, or that span a whole column — using the *extended* form from the 2009
// extension schema: `<x14:dataValidation>` inside `<extLst>`, with the target range in
// `<xm:sqref>`. The legacy writer only understands the standard `<dataValidation>` block,
// so it drops the entire extended block on write and the validation vanishes silently.
//
// Fixture `template.xlsx` (authored in Excel) declares a single list validation over a
// whole column (`A1:A1048576`) entirely in the extended `<extLst>` form — there are zero
// standard `<dataValidation>` entries in it.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'data-validation-extended-survives/template.xlsx';

export default {
  id: 'data-validation-survives-template-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1184},
  cluster: 'core-model',
  description:
    'A data validation a template declares survives a read→write round-trip, including a ' +
    'validation stored in the extended (x14 / extLst) form used for list validations that ' +
    'span whole columns or reference other sheets — it must not be silently dropped on write.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'at least one data validation survives the round-trip in some form',
      baseline: 'pass',
      async expect(api, assert) {
        const {totalValidations} = await api.roundtripFixtureValidationXml(FIXTURE);
        assert.ok(
          totalValidations >= 1,
          `the template's validation must survive write; found ${totalValidations} in the output`
        );
      },
    },
    {
      name: 'the extended (whole-column list) validation is preserved',
      baseline: 'pass',
      async expect(api, assert) {
        const {totalExt, sheets} = await api.roundtripFixtureValidationXml(FIXTURE);
        assert.ok(totalExt >= 1, 'the x14 extended validation should be re-serialized');
        const sqrefs = Object.values(sheets).flatMap(s => s.extSqrefs);
        assert.ok(
          sqrefs.some(ref => /^A1:A104857\d$/.test(ref)),
          `the whole-column target range should survive; got ${JSON.stringify(sqrefs)}`
        );
      },
    },
  ],
};
