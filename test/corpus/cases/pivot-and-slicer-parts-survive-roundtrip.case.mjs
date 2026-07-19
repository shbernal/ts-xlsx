// Cluster: pivot
//
// Real-world scenario: a workbook contains pivot tables, their pivot caches, and slicers (the
// interactive filter widgets that drive pivots). A no-op load→save drops all of it — pivot table
// parts, pivot cache parts, and every slicer part — so the reopened file has lost its pivots and
// slicers. Until a full pivot/slicer model exists, these unmodeled parts must survive a
// round-trip so a fill-and-save workflow does not destroy them.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'pivot-and-slicer-parts-survive-roundtrip/sample.xlsx';

export default {
  id: 'pivot-and-slicer-parts-survive-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2315},
  cluster: 'pivot',
  description:
    'A no-op load→save preserves pivot tables, pivot caches, and slicer parts rather than ' +
    'dropping them — a pivot/slicer-bearing template survives a fill-and-save.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'pivot table and cache parts survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.pivotTables >= 1 && source.pivotCache >= 1, 'precondition: source has pivots');
        assert.strictEqual(rewritten.pivotTables, source.pivotTables, 'pivot table parts survive');
        assert.strictEqual(rewritten.pivotCache, source.pivotCache, 'pivot cache parts survive');
      },
    },
    {
      name: 'slicer parts survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.slicers >= 1, 'precondition: source has slicers');
        assert.strictEqual(rewritten.slicers, source.slicers, `all ${source.slicers} slicer parts survive`);
      },
    },
  ],
};
