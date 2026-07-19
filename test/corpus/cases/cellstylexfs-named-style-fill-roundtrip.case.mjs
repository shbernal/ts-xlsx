// Cluster: styles
//
// Real-world scenario: in the OOXML styles part a cell's effective formatting is layered. A direct
// cellXfs record (the format applied to a cell) may carry an `xfId` pointing into cellStyleXfs — the
// collection of NAMED cell-style definitions — and a visual property such as a fill, font, or border
// can live entirely in that named-style layer rather than being duplicated on the direct cellXfs
// record. Files produced by spreadsheet applications routinely format cells this way (e.g. applying a
// built-in or custom cell style). A reader that ignores the xfId→cellStyleXfs link resolves the cell
// against only its (unstyled) direct format, so the cell renders blank; a writer that discards the
// cellStyleXfs layer drops the named styles entirely, so a load→save round-trip strips the formatting
// and the saved file no longer matches the original.
//
// Fixture `named-style-fill.xlsx` was authored so that A1's yellow fill lives ONLY in a named-style
// (cellStyleXfs) entry: the cell's direct cellXfs xf carries fillId=0 and inherits the fill purely
// through xfId=1 → cellStyleXfs[1] (fillId=2). The fill is invisible unless the named-style layer is
// honored.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'cellstylexfs-named-style-fill/named-style-fill.xlsx';

export default {
  id: 'cellstylexfs-named-style-fill-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A cell whose fill is supplied through a named cell style (a cellXfs xfId into cellStyleXfs, ' +
    'with the fill defined only in the named-style layer) resolves to that fill on read, and a ' +
    'load→save round-trip preserves the cellStyleXfs definitions and the xfId link so the file ' +
    'still renders identically.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the source fixture carries a non-default named-style layer (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {srcCellStyleXfsCount} = await api.namedStyleFillReport(FIXTURE);
        assert.ok(
          srcCellStyleXfsCount >= 2,
          `fixture must declare a named-style entry beyond the default; got cellStyleXfs count ${srcCellStyleXfsCount}`
        );
      },
    },
    {
      name: "a cell's fill defined in the named-style layer is resolved on read",
      baseline: 'pass',
      async expect(api, assert) {
        const {readFill} = await api.namedStyleFillReport(FIXTURE);
        assert.ok(readFill, 'A1 must report a fill inherited through its named cell style, not none');
        assert.strictEqual(readFill.type, 'pattern', 'the resolved fill is the solid pattern from the named style');
        const fg = readFill.fgColor && readFill.fgColor.argb;
        assert.strictEqual(fg, 'FFFFFF00', 'the resolved fill carries the named style yellow');
      },
    },
    {
      name: 'the named-style layer survives a load→save round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {roundtripCellStyleXfsCount} = await api.namedStyleFillReport(FIXTURE);
        assert.ok(
          roundtripCellStyleXfsCount >= 2,
          `the written styles part must retain the named-style (cellStyleXfs) definitions; got count ${roundtripCellStyleXfsCount}`
        );
      },
    },
    {
      name: "the cell's xfId link into the named style survives the round-trip",
      baseline: 'pass',
      async expect(api, assert) {
        const {roundtripCellHasXfIdLink} = await api.namedStyleFillReport(FIXTURE);
        assert.strictEqual(
          roundtripCellHasXfIdLink,
          true,
          'the round-tripped cellXfs entry must retain its xfId link to the named-style record'
        );
      },
    },
  ],
};
