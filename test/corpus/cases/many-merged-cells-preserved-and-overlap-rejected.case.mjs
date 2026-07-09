// Cluster: merged-cells
//
// Real-world scenario: a worksheet has a large number of merged cell ranges (real files reach tens
// of thousands — e.g. one small horizontal merge per row over thousands of rows). Loading and
// re-saving must preserve every merge range exactly, with no merge dropped or duplicated, while
// still rejecting genuinely overlapping merges. (The pathological parse *time* of tens of
// thousands of merges is a separate performance concern captured in a spec note; this case locks
// the correctness of merge preservation on a tractable count.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// One horizontal merge per row across many rows — the shape of the real-world file, at a size the
// corpus can round-trip quickly.
const MERGES = Array.from({length: 60}, (_, i) => `A${i + 1}:B${i + 1}`);
const SPEC = {
  sheets: [{name: 'S', cells: MERGES.map((_, i) => ({ref: `A${i + 1}`, value: i + 1})), merges: MERGES}],
};

export default {
  id: 'many-merged-cells-preserved-and-overlap-rejected',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2689},
  cluster: 'merged-cells',
  description:
    'Every merged range survives a round-trip with identical bounds and count (none dropped or ' +
    'duplicated), while an overlapping merge is still rejected — the overlap check is optimized, ' +
    'not removed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every merged range survives a round-trip with the same count',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.merges.length, MERGES.length, 'no merge dropped or duplicated');
      },
    },
    {
      name: 'each original merge range is present after the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const got = new Set(model.sheets.S.merges);
        for (const m of MERGES) assert.ok(got.has(m), `merge ${m} is preserved`);
      },
    },
    {
      name: 'an overlapping merge is still rejected',
      baseline: 'pass',
      async expect(api, assert) {
        const {error} = await api.mutateWorksheet({
          ops: [{op: 'mergeCells', range: 'A1:B2'}, {op: 'mergeCells', range: 'B2:C3'}],
        });
        assert.ok(error, `overlapping merges must be rejected; got error ${JSON.stringify(error)}`);
      },
    },
  ],
};
