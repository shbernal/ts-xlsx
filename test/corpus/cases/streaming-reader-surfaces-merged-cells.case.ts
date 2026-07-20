// Cluster: streaming
//
// Real-world scenario: a consumer parses a large workbook with the streaming reader (low-memory,
// row-by-row) and needs to know which ranges are merged — to render the sheet, read a value once
// from a merged master, or skip covered cells. The buffered reader exposes merge ranges from the
// worksheet model; the streaming reader drops them entirely (the mergeCells element is never
// surfaced), so a streaming caller cannot recover merge geometry without abandoning streaming and
// loading the whole file — defeating the reason they streamed. The streaming reader must surface the
// same merge set the buffered reader does.
//
// (The design/emission-ordering discussion is captured in the streaming-read-surfaces-merged-cells
// spec note; this case locks the observable gap: streamed merges must equal buffered merges.)

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streaming-reader-surfaces-merged-cells',
  provenance: {source: 'upstream-pr'},
  cluster: 'streaming',
  description:
    'The streaming worksheet reader surfaces the same merged-cell ranges the buffered reader exposes ' +
    'for the same workbook, rather than dropping merge geometry that a low-memory consumer cannot ' +
    'otherwise recover.',

  behavior: [
    {
      name: 'the buffered reader exposes the declared merge ranges (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eagerMerges} = await api.streamReadMergesReport();
        assert.deepStrictEqual(
          eagerMerges,
          ['A1:B2', 'D1:D3'],
          'the buffered reader surfaces both merges',
        );
      },
    },
    {
      name: 'the streaming reader surfaces the same merge ranges',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eagerMerges, streamedMerges, error} = await api.streamReadMergesReport();
        assert.ok(!error, `streaming read must not error; got ${JSON.stringify(error)}`);
        assert.deepStrictEqual(
          streamedMerges,
          eagerMerges,
          `the streaming reader must surface the same merges as the buffered reader; got ${JSON.stringify(streamedMerges)}`,
        );
      },
    },
  ],
} satisfies Case;
