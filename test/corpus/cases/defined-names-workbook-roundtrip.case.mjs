// Cluster: defined-names
//
// Real-world scenario: a user gives a range a memorable name — `TaxRate`, `SalesData` — so formulas
// can reference it by name instead of by coordinates. Excel stores these workbook-level names in a
// `<definedNames>` block inside workbook.xml, each `<definedName name="…">` holding the reference as
// its text content, sited immediately after `<sheets>`. A writer that drops the block, mis-sites it,
// or fails to escape a name/reference produces a workbook whose named references silently vanish.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'defined-names-workbook-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'defined-names',
  description:
    'A workbook-level defined name is persisted in a <definedNames> block with its reference as the ' +
    'element text, and a name or reference carrying XML-special characters is escaped, not corrupted.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a global defined name is written into <definedNames> with its reference as text',
      baseline: 'pass',
      async expect(api, assert) {
        const {definedNames} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', value: 1}]}],
          definedNames: [{name: 'TaxRate', ranges: ['S!$A$1:$B$2']}],
        });
        const entry = definedNames.find(d => d.name === 'TaxRate');
        assert.ok(entry, `TaxRate must appear in <definedNames>; got ${JSON.stringify(definedNames)}`);
        assert.equal(entry.refersTo, 'S!$A$1:$B$2', 'the reference is the element text');
        assert.equal(entry.localSheetId, null, 'a global name carries no localSheetId');
      },
    },
    {
      name: 'a defined name whose reference contains an ampersand is XML-escaped, not corrupted',
      baseline: 'pass',
      async expect(api, assert) {
        const {definedNames} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', value: 1}]}],
          definedNames: [{name: 'Combo', ranges: ["'A & B'!$A$1"]}],
        });
        const entry = definedNames.find(d => d.name === 'Combo');
        assert.ok(entry, `Combo must appear in <definedNames>; got ${JSON.stringify(definedNames)}`);
        assert.ok(entry.refersTo.includes('&amp;'), `the ampersand must be stored as an entity; got ${entry.refersTo}`);
        assert.ok(
          !/&(?!(amp|lt|gt|quot|apos);)/.test(entry.refersTo),
          `the stored text must be XML-well-formed; got ${entry.refersTo}`
        );
      },
    },
  ],
};
