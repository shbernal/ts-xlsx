// Cluster: styles
//
// Real-world scenario: rows and columns can be grouped into a collapsible outline by assigning an
// outline (grouping) level — the OOXML outlineLevel attribute on <row> and <col> (0 = ungrouped, up
// to 7 nested levels). Setting a level on a row and on a column must survive a write/read cycle so the
// grouping is preserved on reopen. Locks basic outline-level round-trip for both axes.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'row-and-column-outline-levels-round-trip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'An outline (grouping) level set on a row and on a column survives a write/read round-trip on ' +
    'its respective axis, so row and column groupings are preserved.',

  behavior: [
    {
      name: 'a row outline level round-trips',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rowOutline} = await api.rowColumnOutlineLevelRoundtrip();
        assert.strictEqual(rowOutline, 1, 'the row keeps outline level 1');
      },
    },
    {
      name: 'a column outline level round-trips',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {colOutline} = await api.rowColumnOutlineLevelRoundtrip();
        assert.strictEqual(colOutline, 1, 'the column keeps outline level 1');
      },
    },
  ],
} satisfies Case;
