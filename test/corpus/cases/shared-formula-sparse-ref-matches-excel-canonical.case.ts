// Cluster: formulas
//
// A shared group whose clones are non-contiguous — master B1, clones B2 and D5 — cannot be described
// by a rectangle that contains only the cloned cells: `planSharedFormulas` emits the bounding box
// `ref="B1:D5"`, fifteen cells of which only three are written. The open question ADR 0012 raised was
// whether that over-wide `ref` is a defect: does a consumer read `ref` as an instruction to
// materialize the translated formula across the whole rectangle? LibreOffice does exactly that; the
// fear was that Excel might too, or might reject the geometry with a repair prompt.
//
// It was seeded through the Excel-oracle harness (ADR 0013) on Excel 16.0 build 20131 — sidecar
// `test/corpus/fixtures/excel-oracle/shared-formula-sparse-ref.json`. The verdict: BENIGN. Excel
// treats the `ref` as a *bounding-box hint*, not an assertion that every enclosed cell is a clone. It
// opened the package without repair, did NOT materialize the empty interior cells, and re-saved a
// byte-structurally identical group (same `ref="B1:D5"`, same `si`, the same two clones). ts-xlsx's
// output is already Excel's own canonical form — so the two candidate "fixes" ADR 0012 floated (split
// into contiguous runs / degrade clones to standalone `<f>`) would make ts-xlsx *diverge* from Excel.
//
// This case is the Tier-2 seam fact that LOCKS that Tier-3 finding in CI (ADR 0012 seed+lock split):
// it reads the emitted geometry straight off the `<f>` elements and asserts the two properties Excel's
// canonical form fixes — one master with exactly `ref="B1:D5"`, and exactly the two authored clones as
// slaves (the interior is never materialized). A regression that split the group or auto-filled the
// interior would break this without ever re-opening Excel.

import type {Assert, Case, CorpusApi} from '../case.ts';

// The probe geometry, mirrored from tools/excel-oracle/probes/shared-formula-sparse-ref.json so the
// locked spec and the seeded observation describe the same workbook.
const SPARSE_GROUP = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A2', value: 2},
        {ref: 'C5', value: 5},
        {ref: 'B1', formula: 'A1*2', result: 2},
        {ref: 'B2', sharedFormula: 'B1', result: 4},
        {ref: 'D5', sharedFormula: 'B1', result: 10},
      ],
    },
  ],
};

export default {
  id: 'shared-formula-sparse-ref-matches-excel-canonical',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/shared-formula-sparse-ref.json',
  },
  cluster: 'formulas',
  description:
    'A non-contiguous shared group (master B1, clones B2 + D5) emits the bounding-box `ref="B1:D5"` ' +
    'with exactly the two authored clones as slaves — the canonical form Excel Desktop itself ' +
    're-saves (ADR 0013). Locks that geometry structurally so a regression that splits the group or ' +
    'materializes the empty interior is caught in CI without re-opening Excel.',

  behavior: [
    {
      // The exact geometry Excel re-saved: a single master over the full bounding box B1:D5. Splitting
      // the group into contiguous runs (the ADR 0012 candidate fix) would produce two masters and a
      // different ref, diverging from Excel's canonical form.
      name: 'the non-contiguous group emits one master with Excel’s canonical bounding-box ref',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {masters} = (await api.inspectPackage(SPARSE_GROUP)).sheets.S.sharedFormulas;
        assert.strictEqual(masters.length, 1, `exactly one master; got ${JSON.stringify(masters)}`);
        assert.strictEqual(masters[0]?.cell, 'B1', 'the master anchors on B1');
        assert.strictEqual(
          masters[0]?.ref,
          'B1:D5',
          `the master ref is Excel's canonical bounding box; got ${JSON.stringify(masters[0])}`,
        );
      },
    },
    {
      // The anti-materialization lock: Excel did NOT auto-fill the eleven empty interior cells, and
      // neither do we — only the two authored clones carry a `t="shared"` slave. A writer that
      // materialized the interior (as LibreOffice does on read) would emit extra slaves here.
      name: 'only the two authored clones are slaves — the empty interior is not materialized',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {slaves, everySlaveHasMaster, everySlaveWithinMasterRef} = (
          await api.inspectPackage(SPARSE_GROUP)
        ).sheets.S.sharedFormulas;
        const slaveCells = slaves.map((s: {cell: string}) => s.cell).sort();
        assert.deepStrictEqual(
          slaveCells,
          ['B2', 'D5'],
          `exactly the two authored clones are slaves; got ${JSON.stringify(slaves)}`,
        );
        assert.strictEqual(
          everySlaveHasMaster,
          true,
          `every slave resolves to the master; got ${JSON.stringify(slaves)}`,
        );
        assert.strictEqual(
          everySlaveWithinMasterRef,
          true,
          `every slave falls inside B1:D5; got ${JSON.stringify(slaves)}`,
        );
      },
    },
  ],
} satisfies Case;
