import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'worksheet-model-preserves-frozen-pane',
  cluster: 'core-model',
  description:
    'A sheet copied through the `model` export/import contract keeps its frozen pane. A frozen ' +
    'view is workbook-independent sheet state: it means the same thing on any sheet of any ' +
    'workbook, so it belongs in the snapshot alongside the merges and the autofilter, and must ' +
    'reach the written package as a `<pane>` rather than being silently unfrozen in transit.',
  provenance: {source: 'adr-0005-amendment', ref: 1},
  behavior: [
    {
      name: 'a model transplant carries the frozen split to the destination sheet',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.frozenPaneSurvivesModelTransplant();
        assert.strictEqual(report.dstState, 'frozen');
        assert.strictEqual(report.dstXSplit, 1);
        assert.strictEqual(report.dstYSplit, 2);
      },
    },
    {
      name: "the destination sheet's written XML carries the frozen `<pane>`",
      expect(api: CorpusApi, assert: Assert) {
        // The in-memory half can hold while the writer drops it, so the package is asserted too:
        // a pane the reader reconstructs from nothing would be a fiction of the round-trip.
        const report = api.frozenPaneSurvivesModelTransplant();
        assert.strictEqual(report.dstPaneEmitted, true);
      },
    },
    {
      name: "reading a sheet's model leaves its own pane intact",
      expect(api: CorpusApi, assert: Assert) {
        const report = api.frozenPaneSurvivesModelTransplant();
        assert.strictEqual(report.srcPaneEmitted, true);
      },
    },
  ],
} satisfies Case;
