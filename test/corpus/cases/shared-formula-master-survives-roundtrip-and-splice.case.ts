// Cluster: formulas
//
// Real-world scenario: a workbook contains shared formulas — the master/slave grouping a spreadsheet
// app emits when a formula is filled across a range (one master cell defining the formula plus
// dependent cells that reference it by a shared range). Reading such a workbook and writing it back,
// or applying a structural edit like inserting a column, must not corrupt the master/slave
// bookkeeping. The writer enforces an invariant that the master cell precede (be above and/or left
// of) each dependent; when a splice shifts the cells, that invariant is violated and the write
// throws "Shared Formula master must exist above and or left of clone". A read/write round-trip works
// today; a column splice into a shared-formula sheet does not.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'shared-formula-master-survives-roundtrip-and-splice',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A shared-formula master/slave group survives a read/write round-trip without throwing and keeps ' +
    'its dependent cells’ formulas; splicing a column into a sheet that holds shared formulas must ' +
    'likewise not throw the "master must exist above/left of clone" error.',

  behavior: [
    {
      name: 'reading a shared-formula workbook and writing it back does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {roundtripOk, roundtripError} = await api.sharedFormulaRoundtripAndSplice();
        assert.strictEqual(
          roundtripOk,
          true,
          `the round-trip must not throw; got ${JSON.stringify(roundtripError)}`,
        );
      },
    },
    {
      name: 'the dependent cells keep a formula after the round-trip (not just a cached value)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {preservedFormulas} = await api.sharedFormulaRoundtripAndSplice();
        assert.strictEqual(
          preservedFormulas,
          true,
          'each dependent cell re-reads as a formula, not a bare value',
        );
      },
    },
    {
      name: 'splicing a column into a shared-formula sheet does not throw the master-position error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {spliceOk, spliceError} = await api.sharedFormulaRoundtripAndSplice();
        assert.strictEqual(
          spliceOk,
          true,
          `a column splice must re-anchor the shared-formula master rather than throwing; got ${JSON.stringify(spliceError)}`,
        );
      },
    },
  ],
} satisfies Case;
