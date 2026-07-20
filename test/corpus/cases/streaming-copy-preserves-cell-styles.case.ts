// Cluster: streaming
//
// Real-world scenario: a user streams a workbook through the streaming reader (with styles cached)
// and rebuilds it through the streaming writer (configured to emit styles), copying each cell's value
// AND style onto the new row. The per-cell font, fill, and number format must survive that streaming
// read→write copy, and the emitted styles part must be well-formed so a strict consumer loads it
// without a styles error.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streaming-copy-preserves-cell-styles',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'Copying a cell’s value and style from the streaming reader onto the streaming writer (with ' +
    'styles enabled) preserves the per-cell font, fill, and number format through the round-trip, ' +
    'and the emitted styles part loads cleanly.',

  behavior: [
    {
      name: 'the streaming style copy completes and the output loads',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {copyError, loadOk} = await api.streamingStyleCopyReport();
        assert.strictEqual(
          copyError,
          null,
          `the streaming copy must not throw; got ${JSON.stringify(copyError)}`,
        );
        assert.strictEqual(loadOk, true, 'the emitted package (and its styles part) loads');
      },
    },
    {
      name: 'the copied cell’s font and number format survive',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {fontBold, fontColor, numFmt} = await api.streamingStyleCopyReport();
        assert.strictEqual(fontBold, true, 'the bold font survives the streaming copy');
        assert.strictEqual(fontColor, 'FFFF0000', 'the font color survives');
        assert.strictEqual(numFmt, '0.00%', 'the number format survives');
      },
    },
    {
      name: 'the copied cell’s fill survives',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasFill} = await api.streamingStyleCopyReport();
        assert.strictEqual(hasFill, true, 'the solid fill survives the streaming copy');
      },
    },
  ],
} satisfies Case;
