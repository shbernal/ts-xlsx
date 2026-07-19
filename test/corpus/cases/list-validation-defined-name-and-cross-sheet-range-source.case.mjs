// Cluster: validations
//
// Real-world scenario: a dropdown (list) data validation is very often driven not by an inline
// literal list but by a *reference* — a workbook defined name pointing at a column on another
// sheet, or an explicit cross-sheet range like `Options!A1:B1`. Both are the standard way to
// keep the allowed values on a separate "options" sheet. Reading such a file must surface each
// validation's source as its reference text (the defined name, or the range) — never stringified
// to "[object Object]" — and a read/write round-trip must preserve that reference and the target
// cell so the dropdown still resolves to the same options.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'list-validation-defined-name-and-cross-sheet-range-source/source.xlsx';

export default {
  id: 'list-validation-defined-name-and-cross-sheet-range-source',
  provenance: {source: 'upstream-issue'},
  cluster: 'validations',
  description:
    'A list-type data validation whose source is a reference — a workbook defined name, or a ' +
    'cross-sheet range like "Options!A1:B1" — is read back as that reference text (not an ' +
    '"[object Object]" string), and both the reference and its target cell survive a round-trip.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a list validation sourced from a cross-sheet range exposes the range as its formula text',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.readFixtureValidationRules(FIXTURE);
        const rules = sheets.Main.rules.map((r) => r.rule);
        const crossSheet = rules.find(
          (r) => r.type === 'list' && (r.formulae || []).some((f) => /^Options!/.test(f)),
        );
        assert.ok(
          crossSheet,
          `expected a list validation sourced from a cross-sheet range; got ${JSON.stringify(rules)}`,
        );
        assert.ok(
          (crossSheet.formulae || []).includes('Options!A1:B1'),
          `the cross-sheet range must be surfaced verbatim, not stringified; got ${JSON.stringify(crossSheet.formulae)}`,
        );
      },
    },
    {
      name: 'a list validation sourced from a defined name exposes the name as its formula text',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.readFixtureValidationRules(FIXTURE);
        const rules = sheets.Main.rules.map((r) => r.rule);
        const named = rules.find(
          (r) => r.type === 'list' && (r.formulae || []).includes('DropdownOptions'),
        );
        assert.ok(
          named,
          `a defined-name list source must be read back as the name "DropdownOptions", not "[object Object]"; got ${JSON.stringify(rules)}`,
        );
      },
    },
    {
      name: 'both reference-based list validations survive a read/write round-trip as standard validations',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripFixtureValidationXml(FIXTURE);
        const rules = Object.values(sheets).flatMap((s) => s.standardRules);
        const sources = rules.filter((r) => r.type === 'list').map((r) => r.formula1);
        assert.ok(
          sources.includes('Options!A1:B1'),
          `the cross-sheet range source must survive re-serialization; got ${JSON.stringify(sources)}`,
        );
        assert.ok(
          sources.includes('DropdownOptions'),
          `the defined-name source must survive re-serialization; got ${JSON.stringify(sources)}`,
        );
      },
    },
  ],
};
