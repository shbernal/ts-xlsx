// Cluster: tables
//
// Real-world scenario: a worksheet defines several adjacent columns that share the same width,
// styling, and outline level. On write, the serializer coalesces runs of equivalent columns into
// shared <col min="…" max="…"> spans (rather than emitting one <col> per column). That collapse pass
// compares neighbouring column definitions for equivalence; it must complete without throwing, and
// the produced package must reload. A reported crash in this pass ("column.equivalentTo is not a
// function" while building the worksheet model) must not occur for ordinary defined columns.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'equivalent-columns-collapse-on-write-without-crash',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A worksheet with several adjacent equivalent columns (same width and outline level) writes ' +
    'without crashing during the equivalent-column collapse pass, coalesces them into shared <col> ' +
    'spans, and reloads cleanly.',

  behavior: [
    {
      name: 'writing equivalent adjacent columns does not throw and the package reloads',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, writeError, reloadOk} = await api.equivalentColumnCollapseReport();
        assert.strictEqual(
          writeOk,
          true,
          `the write must not throw; got ${JSON.stringify(writeError)}`,
        );
        assert.strictEqual(reloadOk, true, 'the written package reloads');
      },
    },
    {
      name: 'equivalent adjacent columns are coalesced into fewer <col> spans than columns',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {colSpanCount} = await api.equivalentColumnCollapseReport();
        assert.ok(
          colSpanCount >= 1 && colSpanCount < 4,
          `four equivalent columns should collapse into fewer than four <col> spans; got ${colSpanCount}`,
        );
      },
    },
  ],
} satisfies Case;
