// Cluster: security
//
// Real-world scenario: a worksheet part declares `<col min="1" max="99999999"/>`. The span names
// more columns than the format has, and a reader that walks it verbatim materializes one column
// record per index -- roughly 16.7 million of them -- spinning for tens of seconds before dying on
// the map's size limit. That is a denial of service from a one-line input, so the span must be
// bounded by the format's own ceiling rather than by the file's arithmetic. The same question on
// the other axis: a `<row r>` past the last row names a position that does not exist, and must not
// enter the model either.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'hostile-column-span-is-bounded-on-read',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A <col> span or <row r> naming a position outside the spreadsheet grid is bounded on read: ' +
    'the span is clamped to the last real column, an element wholly outside the grid is dropped, ' +
    'and neither can drive an unbounded allocation.',

  behavior: [
    {
      name: 'a <col> span running past the last column is clamped to it, not walked',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.hostileGridSpanReport();
        assert.equal(
          report.spanColumnCount,
          report.maxColumn,
          'the span may create at most one record per real column',
        );
        assert.equal(
          report.lastSpannedWidth,
          12,
          'and it still reaches the last real column, so the clamp loses no legal content',
        );
      },
    },
    {
      name: 'a <col> or <row> wholly outside the grid puts nothing into the model',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.hostileGridSpanReport();
        assert.equal(report.beyondColumnCount, 0, 'no column past the last one is recorded');
        assert.equal(report.beyondRowCount, 0, 'no row past the last one is recorded');
      },
    },
  ],
} satisfies Case;
