// Cluster: images
//
// Real-world scenario: an image is placed with a two-cell anchor whose top-left coordinate has a
// fractional component (e.g. col 0.5, row 0.5) — the anchor point sits halfway into a cell. When
// the target column widths and row heights differ from the defaults, the fractional part must be
// converted to an EMU offset scaled by that cell's ACTUAL width and height: a col-0.5 anchor in a
// wide column must sit further right than the same anchor in a narrow column, and a row-0.5 anchor
// in a tall row further down than in a short row. A conversion that used the default cell size
// regardless of the real dimensions would render the image off-center.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const fractionalAnchorSheet = ({colWidth, rowHeight}) => ({
  sheets: [
    {
      name: 'S',
      columns: colWidth === undefined ? [] : [{index: 1, width: colWidth}],
      rows: rowHeight === undefined ? [] : [{index: 1, height: rowHeight}],
      images: [{range: {tl: {col: 0.5, row: 0.5}, br: {col: 2.5, row: 2.5}}}],
    },
  ],
});

const firstFrom = async (api, spec) => {
  const {anchors} = await api.inspectImageAnchors(spec);
  return anchors[0].from;
};

export default {
  id: 'image-anchor-fractional-offset-respects-cell-size',
  provenance: {source: 'upstream-pr'},
  cluster: 'images',
  description:
    'The sub-cell EMU offset of a fractional image anchor scales with the target cell\'s real ' +
    'column width and row height — a col-0.5 anchor sits further right in a wide column, and a ' +
    'row-0.5 anchor further down in a tall row, rather than using the default cell size.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a fractional column anchor offset grows with the column width',
      baseline: 'pass',
      async expect(api, assert) {
        const narrow = await firstFrom(api, fractionalAnchorSheet({colWidth: 5}));
        const wide = await firstFrom(api, fractionalAnchorSheet({colWidth: 50}));
        assert.ok(narrow.colOff > 0 && wide.colOff > 0, 'both offsets are positive sub-cell offsets');
        assert.ok(
          wide.colOff > narrow.colOff,
          `a wider column must produce a larger sub-cell colOff; narrow=${narrow.colOff} wide=${wide.colOff}`
        );
      },
    },
    {
      name: 'a fractional row anchor offset grows with the row height',
      baseline: 'pass',
      async expect(api, assert) {
        const short = await firstFrom(api, fractionalAnchorSheet({rowHeight: 10}));
        const tall = await firstFrom(api, fractionalAnchorSheet({rowHeight: 80}));
        assert.ok(short.rowOff > 0 && tall.rowOff > 0, 'both offsets are positive sub-cell offsets');
        assert.ok(
          tall.rowOff > short.rowOff,
          `a taller row must produce a larger sub-cell rowOff; short=${short.rowOff} tall=${tall.rowOff}`
        );
      },
    },
  ],
};
