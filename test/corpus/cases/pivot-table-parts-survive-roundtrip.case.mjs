// Cluster: pivot
//
// Real-world scenario: a workbook contains pivot tables backed by pivot-table parts and their
// pivot caches (definition + records). A library that does not model pivots drops all of it on a
// no-op load→save: the pivot table parts and the pivot cache vanish, so the reopened file has no
// pivots. Until a full pivot model exists, these unmodeled parts must at least survive a
// round-trip so a "load a template, fill cells, save" workflow does not destroy the user's pivots.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'pivot-table-parts-survive-roundtrip/sample.xlsx';

export default {
  id: 'pivot-table-parts-survive-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1678},
  cluster: 'pivot',
  description:
    'A no-op load→save preserves pivot-table parts and their pivot caches instead of dropping ' +
    'them — a chart/pivot-bearing template survives a fill-and-save without losing its pivots.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'pivot table parts survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.pivotTables >= 1, 'precondition: source has pivot tables');
        assert.strictEqual(
          rewritten.pivotTables,
          source.pivotTables,
          `all ${source.pivotTables} pivot table parts survive`,
        );
      },
    },
    {
      name: 'pivot cache parts survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.pivotCache >= 1, 'precondition: source has pivot cache parts');
        assert.strictEqual(
          rewritten.pivotCache,
          source.pivotCache,
          `all ${source.pivotCache} pivot cache parts survive`,
        );
      },
    },
  ],
};
