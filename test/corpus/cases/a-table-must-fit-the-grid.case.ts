// Cluster: tables
//
// Real-world scenario: a table is declared by an anchor plus a column list and a data-row count, and
// the rectangle it occupies is derived from those. The anchor was checked against the grid and the
// far corner was not, so a table could be declared over columns past XFD or rows past 1048576 -- a
// rectangle Excel has no room for. That does not fail loudly. `range`, `autoFilterRef` and `region`
// all throw when anything reads the table back, so the sheet becomes un-serialisable and
// un-inspectable; and where the count is small enough to encode, the writer emits a `<table ref>`
// naming rows that cannot exist, which is a package Excel meets with its repair prompt. The row
// count reaches this from a file as well as from an author: a reader derives it from the stored
// range. The same corner moves under a column splice, where clamping the anchor is not enough: an
// anchor clamped onto XFD still puts a two-column table's right edge at XFE.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-table-must-fit-the-grid',
  provenance: {source: 'grid-edit-audit'},
  cluster: 'tables',
  description:
    'A table occupies a rectangle derived from its anchor, its column count and its row count. ' +
    'That whole rectangle must lie inside the grid: a table reaching past the last column or the ' +
    'last row is refused when it is declared, and a column splice that would push its right edge ' +
    'past XFD drops it rather than producing one nothing can read back.',

  behavior: [
    {
      name: 'a table whose columns reach past the last one is refused, naming the table',
      async expect(api: CorpusApi, assert: Assert) {
        const {pastRightEdge} = await api.tableOutsideTheGrid();
        assert.ok(
          pastRightEdge !== null,
          'three columns anchored at XFC reach XFE, which the grid does not have',
        );
        assert.match(pastRightEdge, /table "T"/, 'the message names the table, not a bare column');
        assert.match(pastRightEdge, /XFD/, 'and says what the last column is');
      },
    },
    {
      name: 'a table whose rows reach past the last one is refused too',
      async expect(api: CorpusApi, assert: Assert) {
        const {pastBottomEdge} = await api.tableOutsideTheGrid();
        assert.ok(pastBottomEdge !== null, 'five million data rows do not fit in 1048576');
        assert.match(pastBottomEdge, /table "T"/);
        assert.match(pastBottomEdge, /1048576/);
      },
    },
    {
      name: 'a table ending exactly on the last column is accepted and written as itself',
      async expect(api: CorpusApi, assert: Assert) {
        const {onTheEdge, writtenRef} = await api.tableOutsideTheGrid();
        assert.strictEqual(onTheEdge, 'XFC1:XFD2', 'the last legal rectangle is legal');
        assert.strictEqual(writtenRef, 'XFC1:XFD2', 'and it reaches the package unchanged');
      },
    },
    {
      name: 'a column splice with no room left for the table drops it instead of unreading it',
      async expect(api: CorpusApi, assert: Assert) {
        const {splicedPastRightEdge} = await api.tableOutsideTheGrid();
        assert.strictEqual(
          splicedPastRightEdge.tableCount,
          0,
          'the answer a row splice already gives a table with no room left for its rows',
        );
        assert.strictEqual(
          splicedPastRightEdge.range,
          null,
          'no table survives, so nothing throws on read; a message here is the range that did',
        );
      },
    },
  ],
} satisfies Case;
