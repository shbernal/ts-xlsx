// Cluster: images
//
// Real-world scenario: to place a picture at a precise sub-cell position, a caller anchors
// it with a *fractional* cell coordinate — `tl.col = 3.5` meaning "halfway across the
// fourth column". The library must translate that fraction into an OOXML sub-cell offset
// (`<xdr:colOff>` in EMU) computed against the *actual* width of that column. If the
// column is wide, halfway across it is a large offset; if narrow, a small one.
//
// The bug: the offset is computed from an ad-hoc width→EMU factor that does not match a
// column's real geometry, and default-width columns use a fixed constant. The result is
// that a *wider* custom column can yield a *smaller* half-way offset than a default column
// — the picture lands in the wrong place, and worse the wrong way round. A whole-integer
// anchor (no fraction) has always sat exactly on the cell boundary and must keep doing so.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'fractional-image-anchor-positioning',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 894},
  cluster: 'images',
  description:
    'A fractional image anchor coordinate (e.g. col 3.5) must map to a sub-cell EMU offset ' +
    'proportional to the referenced column’s actual width, so halfway across a wide column ' +
    'is a larger offset than halfway across a narrow default column; a whole-integer anchor ' +
    'sits exactly on the cell boundary with zero offset.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a whole-integer anchor sits exactly on the cell boundary (zero sub-cell offset)',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors({
          sheets: [{name: 'S', images: [{range: {tl: {col: 2, row: 3}, br: {col: 4, row: 6}}}]}],
        });
        const a = anchors[0];
        assert.ok(a, 'an image anchor must be serialized');
        assert.strictEqual(a.from.col, 2, 'from column is the integer coordinate');
        assert.strictEqual(a.from.colOff, 0, 'no sub-cell column offset for an integer anchor');
        assert.strictEqual(a.from.rowOff, 0, 'no sub-cell row offset for an integer anchor');
      },
    },
    {
      name: 'halfway across a wide custom column is a larger offset than halfway across a default column',
      baseline: 'pass',
      async expect(api, assert) {
        const wide = await api.inspectImageAnchors({
          sheets: [{name: 'S', columns: [{index: 4, width: 38}], images: [{range: {tl: {col: 3.5, row: 1}, br: {col: 4.5, row: 2}}}]}],
        });
        const def = await api.inspectImageAnchors({
          sheets: [{name: 'S', images: [{range: {tl: {col: 3.5, row: 1}, br: {col: 4.5, row: 2}}}]}],
        });
        const wideOff = wide.anchors[0].from.colOff;
        const defOff = def.anchors[0].from.colOff;
        assert.ok(
          wideOff >= defOff,
          `halfway across a width-38 column (${wideOff} EMU) must be at least halfway across a ` +
            `default-width column (${defOff} EMU); a wider column must not produce a smaller offset`
        );
      },
    },
    {
      name: 'a fractional anchor floors to the containing cell and carries a non-zero sub-cell offset',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors({
          sheets: [{name: 'S', images: [{range: {tl: {col: 3.5, row: 1.5}, br: {col: 4.5, row: 2.5}}}]}],
        });
        const a = anchors[0];
        assert.strictEqual(a.from.col, 3, 'col 3.5 floors to cell column 3');
        assert.strictEqual(a.from.row, 1, 'row 1.5 floors to cell row 1');
        assert.ok(a.from.colOff > 0, 'the .5 fraction produces a non-zero column offset');
        assert.ok(a.from.rowOff > 0, 'the .5 fraction produces a non-zero row offset');
      },
    },
  ],
};
