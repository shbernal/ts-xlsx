// Cluster: merges
//
// Real-world scenario: a worksheet ends with a row that is both empty (no explicit cell values) and
// covered by a merged range whose master lies above it (A1:B3 with values only in A1/A2). When
// iterating row-by-row then cell-by-cell (including empty cells), every covered position of the
// merged region must be reachable, and the leading cell of the trailing merged row (A3) must be
// visited and resolve to its master. This locks complete iteration and merge metadata for a merged
// range extending into an otherwise-empty final row, so consumers building a merge map see every cell.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'trailing-empty-merged-row-iteration-visits-all-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'merges',
  description:
    'A merged range extending into a trailing, otherwise-empty final row round-trips, and iterating ' +
    'the worksheet (including empty cells) visits the leading cell of that row, which resolves to ' +
    'its master rather than being skipped.',

  behavior: [
    {
      name: 'the trailing merged row is within the sheet bounds after round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rowCount} = await api.trailingMergedRowIterationReport();
        assert.ok(
          rowCount >= 3,
          `the merged range's last row (3) is included in the bounds; got ${rowCount}`,
        );
      },
    },
    {
      name: 'iterating cells visits the leading cell of the trailing merged row',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {a3} = await api.trailingMergedRowIterationReport();
        assert.strictEqual(
          a3.visited,
          true,
          'A3 (leading cell of the trailing merged row) is visited',
        );
      },
    },
    {
      name: 'the leading cell of the trailing merged row resolves to its master',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {a3} = await api.trailingMergedRowIterationReport();
        assert.strictEqual(a3.isMerged, true, 'A3 reports being merged');
        assert.strictEqual(a3.master, 'A1', 'A3 resolves to the merge master A1');
      },
    },
  ],
} satisfies Case;
