import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'sheet-images-carry-between-workbooks',
  cluster: 'images',
  description:
    'Consolidating several files into one workbook means copying image-bearing sheets between ' +
    'workbooks. An image is a media+relationship pair, not a value that travels with a cell: the ' +
    "anchor holds an id into its own workbook's media registry, and that id names a different " +
    'picture, or none, in the destination. Carrying a sheet must therefore register the picture in ' +
    'the destination workbook and rebind the anchor to the id it lands on, and an anchor that ' +
    'reaches the writer still holding a foreign id must be refused by name rather than emitted as ' +
    'a drawing relationship pointing at media that was never written.',
  provenance: {
    source: 'backlog-spec',
    ref: 'image-carry-between-workbooks-requires-media-registration',
  },
  behavior: [
    {
      name: "a carried sheet's picture is registered in the destination package",
      expect(api: CorpusApi, assert: Assert) {
        const report = api.carrySheetImagesAcrossWorkbooks();
        assert.strictEqual(report.dstMediaCount, 2, 'the anchored picture and the background');
      },
    },
    {
      name: 'every drawing embed in the destination resolves to a real media part',
      expect(api: CorpusApi, assert: Assert) {
        // The failure this exists to catch renders as a broken image in Excel rather than an
        // error, so the relationship graph is asserted directly instead of the write succeeding.
        const report = api.carrySheetImagesAcrossWorkbooks();
        assert.strictEqual(report.allEmbedsResolve, true);
      },
    },
    {
      name: 'reading the destination back surfaces the image anchored where it was placed',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.carrySheetImagesAcrossWorkbooks();
        assert.strictEqual(report.reReadImageCount, 1);
        assert.deepEqual(report.reReadAnchor, {col: 1, row: 1});
        assert.strictEqual(report.reReadHasBackground, true);
      },
    },
    {
      name: 'carrying a sheet copies its pictures rather than moving them off the source',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.carrySheetImagesAcrossWorkbooks();
        assert.strictEqual(report.srcStillShowsImage, true);
      },
    },
    {
      name: 'an anchor holding an unregistered media id is refused by name, not written dangling',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.carrySheetImagesAcrossWorkbooks();
        assert.match(String(report.danglingAnchorError), /image id 41/);
        assert.match(String(report.danglingAnchorError), /not registered on the workbook/);
      },
    },
  ],
} satisfies Case;
