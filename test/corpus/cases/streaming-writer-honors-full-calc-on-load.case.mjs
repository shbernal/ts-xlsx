// Cluster: streaming
//
// Real-world scenario: a workbook is produced through the streaming writer and contains formula cells
// whose cached results the producer cannot compute (it is not a formula engine). The standard OOXML
// remedy is to set the workbook's fullCalcOnLoad flag so the consuming spreadsheet app recalculates
// every formula when it opens the file. The in-memory writer honors this — it emits
// <calcPr … fullCalcOnLoad="1"/>. The streaming writer does not: the flag set on the writer never
// reaches the output, whose <calcPr> carries only a calcId, so the produced file opens showing stale
// cached results with no recalculation. Recalc-on-load must work identically on both writers.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-writer-honors-full-calc-on-load',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'Setting fullCalcOnLoad (recalc-on-load) on the streaming writer emits it in the workbook ' +
    'calcPr, matching the in-memory writer, so a streamed workbook tells the consuming app to ' +
    'recalculate on open; with the flag unset, neither writer emits it.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the in-memory writer emits fullCalcOnLoad when set (parity oracle)',
      baseline: 'pass',
      async expect(api, assert) {
        const {memoryHasFlag} = await api.streamingFullCalcOnLoadReport();
        assert.strictEqual(memoryHasFlag, true, 'the in-memory writer serializes fullCalcOnLoad="1"');
      },
    },
    {
      name: 'the streaming writer emits fullCalcOnLoad when set',
      baseline: 'pass',
      async expect(api, assert) {
        const {streamHasFlag, streamSetThrew} = await api.streamingFullCalcOnLoadReport();
        assert.strictEqual(
          streamHasFlag,
          true,
          `the streamed workbook must carry fullCalcOnLoad in its calcPr (setThrew=${streamSetThrew})`
        );
      },
    },
    {
      name: 'the streaming writer omits fullCalcOnLoad when it is not set (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {streamDefaultHasFlag} = await api.streamingFullCalcOnLoadReport();
        assert.strictEqual(streamDefaultHasFlag, false, 'with the flag unset, no fullCalcOnLoad is emitted');
      },
    },
  ],
};
