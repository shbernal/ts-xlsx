// Cluster: styles
//
// Shading a header band, boxing a column of figures, centring a title row — the everyday formatting
// jobs are all rectangular, and the only way to say them used to be a nested loop over rows and
// columns. That loop is easy to get subtly wrong in two directions at once.
//
// **It skips the empty cells.** A block is styled for how it *looks*, and an empty cell in the middle
// of a shaded band still has to be shaded — which means it has to exist, as a styled-but-valueless
// cell. A loop written over the populated cells leaves holes in the band; one written over every
// address has to know to materialise.
//
// **It can mint a style per cell.** The styles part is a shared table referenced by index, and a
// uniformly styled block must collapse to one entry. Getting that wrong is invisible in the rendered
// file and shows up only as a package that takes forever to write and megabytes to store.
//
// A range handle has to make both correct by construction, and keep composing with everything else:
// styling a block must not clear formatting the cells already carried, and a per-cell edit afterwards
// must still win inside the block.

import type {Assert, Case, CorpusApi} from '../case.ts';

const HEADER_FILL = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFBB2649'}};

export default {
  id: 'range-styling-covers-empty-cells-and-still-dedups',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'Styling a rectangular block in one call reaches every cell it covers — including cells that ' +
    'were empty, which are materialised as styled-but-valueless so the band has no holes — while a ' +
    'uniformly styled block still collapses to a single shared style-table entry rather than one ' +
    'per cell.',

  behavior: [
    {
      name: 'an empty cell inside a styled block is materialised and renders styled',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.rangeStyleReport({
          values: [{address: 'A1', value: 'Region'}],
          blocks: [{ref: 'A1:D1', facet: 'fill', value: HEADER_FILL}],
          probe: {ref: 'A1:D1', facet: 'fill', outside: 'E1'},
        });
        assert.strictEqual(report.materialisedInBlock, 4, 'all four positions exist');
        assert.strictEqual(report.styledCells, 4, 'and all four are written with a style');
        // The three that were empty when the block was styled still carry the fill after a
        // round-trip — the band has no holes in it. (The reloaded fill also carries the automatic
        // `bgColor` the writer forces onto a solid pattern, which is why the colour is checked
        // rather than the whole record.)
        for (const address of ['A1', 'B1', 'C1', 'D1']) {
          assert.deepStrictEqual(
            report.reloadedStyles[address]?.fgColor,
            {argb: 'FFBB2649'},
            address,
          );
          assert.strictEqual(report.reloadedStyles[address]?.pattern, 'solid', address);
        }
        assert.strictEqual(report.outsideBlock, null, 'and nothing outside the block was touched');
      },
    },
    {
      name: 'a uniformly styled block collapses to one shared style entry',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // A 32-row band. One entry, not 32: the styles part is a shared table indexed by the cell's
        // `s`, and minting an entry per cell is the historical performance cliff.
        const report = api.rangeStyleReport({
          blocks: [
            {
              ref: 'A2:A33',
              style: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}},
            },
          ],
        });
        assert.strictEqual(report.styledCells, 32);
        assert.strictEqual(report.distinctStyleIds, 1);
        // The default xf plus exactly one more.
        assert.strictEqual(report.cellXfs, 2);
      },
    },
    {
      name: 'two differently styled blocks intern separately, without over-collapsing',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.rangeStyleReport({
          blocks: [
            {ref: 'A1:D1', facet: 'fill', value: HEADER_FILL},
            {ref: 'A2:A33', facet: 'font', value: {bold: true}},
          ],
        });
        assert.strictEqual(report.styledCells, 36);
        assert.strictEqual(report.distinctStyleIds, 2, 'dedup collapses alike, never unalike');
      },
    },
    {
      name: 'styling a block composes with the formatting its cells already carried',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // Setting a fill must not wipe a font, a border or a number format already there — the
        // failure mode of "assign the whole style" applied blindly across a block.
        const report = api.rangeStyleReport({
          blocks: [
            {ref: 'A1:B2', facet: 'font', value: {bold: true}},
            {ref: 'A1:B2', style: {alignment: {horizontal: 'center'}}},
          ],
          probe: {ref: 'A1:B2', facet: 'font', outside: 'C1'},
        });
        assert.deepStrictEqual(report.sharedFacet, {bold: true}, 'the earlier facet survived');
        assert.strictEqual(report.distinctStyleIds, 1);
      },
    },
    {
      name: 'a block reports a facet only when every cell it covers agrees',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // Reading a block that was never styled as a whole must not pick a corner's answer and pass
        // it off as the block's — the holes render unstyled, so the block does not have that facet.
        const report = api.rangeStyleReport({
          blocks: [{ref: 'A1:A1', facet: 'fill', value: HEADER_FILL}],
          probe: {ref: 'A1:D1', facet: 'fill', outside: 'E1'},
        });
        assert.strictEqual(report.sharedFacet, null);
      },
    },
  ],
} satisfies Case;
