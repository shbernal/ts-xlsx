// Cluster: merged-cells
//
// Real-world scenario: a user builds a two-row header by merging a horizontal span (e.g. B1:G1 for a
// group title over the sub-headers beneath), gives the anchor cell a value and center alignment, and
// writes. A reported failure was that Excel showed the "we found a problem… recover?" repair dialog,
// caused by the covered (non-anchor) cells of the merge being emitted as conflicting populated
// content. A clean merge declares the range exactly once and emits a value only on the top-left
// anchor; the covered cells carry no conflicting value, so the file opens without a repair prompt and
// the anchor's value and alignment round-trip.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'merged-range-opens-without-repair-prompt',
  provenance: {source: 'upstream-issue'},
  cluster: 'merged-cells',
  description:
    'Merging a horizontal range with a value + alignment on the anchor emits the merge exactly once ' +
    'with no populated covered cells, so the package opens without a repair prompt, and the anchor’s ' +
    'value and alignment survive a round-trip.',

  behavior: [
    {
      name: 'the merged range is declared exactly once',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {mergeCount, merges} = await api.mergeCleanReport();
        assert.strictEqual(
          mergeCount,
          1,
          `exactly one mergeCell entry; got ${JSON.stringify(merges)}`,
        );
      },
    },
    {
      name: 'the covered non-anchor cells are not emitted as populated content',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {populatedCoveredCells} = await api.mergeCleanReport();
        assert.deepStrictEqual(
          populatedCoveredCells,
          [],
          `covered cells must carry no conflicting value; got ${JSON.stringify(populatedCoveredCells)}`,
        );
      },
    },
    {
      name: 'the anchor cell’s value and alignment survive a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchorValue, anchorAlignment} = await api.mergeCleanReport();
        assert.strictEqual(anchorValue, 'Group Title', 'the anchor value round-trips');
        assert.strictEqual(
          anchorAlignment?.horizontal,
          'center',
          'the anchor alignment round-trips',
        );
      },
    },
  ],
} satisfies Case;
