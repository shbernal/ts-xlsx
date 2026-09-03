// Cluster: security
//
// Real-world scenario: clamping a single `<col min max>` span to the last real column is the fix a
// reader reaches for first, and it is only half the guard. It bounds one element's loop; nothing
// bounds the number of `<col>` elements, and each may span the whole grid, so the cost of a sheet is
// their product. 154 KB of worksheet XML holding two thousand full-grid spans took the buffered
// reader 13 seconds and the streaming one 1.7, which against the package inflate ceiling extrapolates
// to hours of CPU from a
// file that compresses to a few hundred KB. The zip-bomb guard cannot see this one, because after
// inflation the payload genuinely is small: it is a CPU bomb, not an allocation one.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Both readers do this input in under 150 ms once bounded, and both were over 3 s before, so the
// budget sits an order of magnitude clear of each side: it measures the bound rather than the machine.
const TIME_BUDGET_MS = 1_500;

export default {
  id: 'repeated-full-grid-column-spans-are-time-bounded',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A worksheet declaring thousands of full-grid <col> spans is read in bounded time by both the ' +
    'buffered and the streaming reader: the per-sheet budget bounds the total column work, not ' +
    'merely the reach of any one span.',

  behavior: [
    {
      name: 'the buffered reader finishes a sheet of repeated full-grid spans within its budget',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.repeatedFullGridColumnSpanReport();
        assert.ok(
          report.bufferedMs < TIME_BUDGET_MS,
          `${report.spans} full-grid spans in ${report.xmlBytes} bytes of XML took ` +
            `${report.bufferedMs.toFixed(0)}ms to read`,
        );
      },
    },
    {
      name: 'the streaming reader is bounded on the same terms',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.repeatedFullGridColumnSpanReport();
        assert.ok(
          report.streamingMs < TIME_BUDGET_MS,
          `streaming the same sheet took ${report.streamingMs.toFixed(0)}ms`,
        );
      },
    },
    {
      name: 'and the bound costs no legal content: the grid is still fully described',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.repeatedFullGridColumnSpanReport();
        assert.equal(
          report.columnCount,
          report.maxColumn,
          'the first span still reaches every real column',
        );
      },
    },
  ],
} satisfies Case;
