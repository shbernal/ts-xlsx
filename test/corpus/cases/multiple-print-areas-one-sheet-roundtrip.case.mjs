// Cluster: tables
//
// Real-world scenario: Excel lets a single worksheet declare more than one print area (disjoint
// blocks that each print on their own page). In the package this is ONE defined name of type
// _xlnm.Print_Area whose value is a comma-separated list of ranges sharing one localSheetId — e.g.
// "Sheet!$A$1:$F$10,Sheet!$A$12:$F$21". A reader that looks only at the first range silently drops
// the rest, so a two-print-area file comes back with one. Symmetrically, authoring two print areas
// must emit both back into a single Print_Area name as a comma-separated list; splitting on the
// wrong separator or writing only the first corrupts the output.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'multiple-print-areas-one-sheet-roundtrip/source.xlsx';

export default {
  id: 'multiple-print-areas-one-sheet-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A worksheet with two disjoint print areas (recorded as one comma-separated _xlnm.Print_Area ' +
    'defined name) is read back with BOTH ranges and re-emitted with both — not truncated to the ' +
    'first range on read, nor mangled to a single range on write.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the source file declares two print-area ranges in one Print_Area name (oracle)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sourceRangeCount} = await api.roundtripFixturePrintAreas(FIXTURE);
        assert.strictEqual(sourceRangeCount, 2, 'the fixture declares two print areas');
      },
    },
    {
      name: 'reading recovers both print-area ranges, not just the first',
      baseline: 'pass',
      async expect(api, assert) {
        const {readPrintArea} = await api.roundtripFixturePrintAreas(FIXTURE);
        const rangeCount = String(readPrintArea || '')
          .split(',')
          .filter(Boolean).length;
        assert.ok(
          rangeCount >= 2,
          `both print areas must be recovered on read; got printArea=${JSON.stringify(readPrintArea)}`,
        );
      },
    },
    {
      name: 'writing the file back preserves both print-area ranges in the Print_Area name',
      baseline: 'pass',
      async expect(api, assert) {
        const {rewrittenRangeCount} = await api.roundtripFixturePrintAreas(FIXTURE);
        assert.strictEqual(rewrittenRangeCount, 2, 'both ranges must survive re-serialization');
      },
    },
    {
      name: 'authoring two print areas emits both ranges in one sheet-scoped Print_Area name',
      baseline: 'pass',
      async expect(api, assert) {
        const {ranges} = await api.writePrintAreaDefinedName('A1:F10,A12:F21');
        // Both emitted entries must be proper rectangular ranges (a "$A$1:$F$10" shape). A mangled
        // write drops the second range's tail, leaving a bare "A12" that is not a range at all.
        const rectangles = ranges.filter((r) => /:/.test(r) && /\$?[A-Z]+\$?\d+/.test(r));
        assert.strictEqual(
          rectangles.length,
          2,
          `a two-range printArea must emit two proper rectangular ranges, not a truncated/mangled one; got ${JSON.stringify(ranges)}`,
        );
      },
    },
  ],
};
