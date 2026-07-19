// Cluster: styles
//
// Real-world scenario: a workbook has a conditional-formatting rule of type "duplicateValues" over
// a column range, referencing a differential format (dxfId) at a given priority. Read the file and
// write it straight back and the rule is dropped — the emitted worksheet loses the cfRule (the
// whole conditionalFormatting block vanishes, or is reduced to an empty shell). Excel then reports
// the file as damaged and discards the formatting on repair. Even a rule type the library does not
// otherwise interpret must be preserved across a round-trip for input/output fidelity. This is one
// instance of the broader problem: conditional-formatting rule types the library does not model are
// silently lost on save.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'conditional-formatting-duplicate-values-roundtrip/source.xlsx';

export default {
  id: 'conditional-formatting-duplicate-values-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A "duplicateValues" conditional-formatting rule (with its dxfId and priority) survives a ' +
    'read/write round-trip — the writer must not drop the cfRule or emit an empty ' +
    'conditionalFormatting shell, which corrupts the file.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the source file declares a duplicateValues cfRule (oracle)',
      baseline: 'pass',
      async expect(api, assert) {
        const {source} = await api.roundtripFixtureConditionalFormatting(FIXTURE);
        const dup = source.rules.find((r) => r.type === 'duplicateValues');
        assert.ok(
          dup,
          `the fixture must declare a duplicateValues rule; got ${JSON.stringify(source)}`,
        );
        assert.strictEqual(dup.dxfId, '0', 'the source rule references dxfId 0');
      },
    },
    {
      name: 'the duplicateValues rule survives a no-op round-trip with its type, dxfId, and priority',
      baseline: 'pass',
      async expect(api, assert) {
        const {rewritten} = await api.roundtripFixtureConditionalFormatting(FIXTURE);
        const dup = rewritten.rules.find((r) => r.type === 'duplicateValues');
        assert.ok(
          dup,
          `the duplicateValues rule must survive re-serialization, not be dropped; rewritten=${JSON.stringify(rewritten)}`,
        );
        assert.strictEqual(dup.dxfId, '0', 'the rewritten rule keeps its dxfId reference');
        assert.ok(dup.priority != null, 'the rewritten rule keeps a priority');
      },
    },
    {
      name: 'the re-written worksheet has no empty conditionalFormatting shell (block without a rule)',
      baseline: 'pass',
      async expect(api, assert) {
        const {rewritten} = await api.roundtripFixtureConditionalFormatting(FIXTURE);
        assert.ok(
          rewritten.rules.length >= rewritten.blockCount,
          `every conditionalFormatting block must still contain at least one cfRule; blocks=${rewritten.blockCount} rules=${rewritten.rules.length}`,
        );
      },
    },
  ],
};
