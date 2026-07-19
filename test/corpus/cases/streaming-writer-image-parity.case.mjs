// Cluster: images
//
// Real-world scenario: a caller builds a large spreadsheet through the streaming/incremental writer
// — chosen because the dataset is too big to hold fully in memory — and wants to embed an image
// (a logo anchored over a cell range) exactly as they would with the in-memory workbook. In the
// in-memory path they register the image on the workbook and anchor it onto a worksheet over a
// range. On the streaming path the equivalent is absent: the streamed worksheet exposes no way to
// anchor a registered image, so a caller who needs both large data AND an image is forced back onto
// the in-memory path they could not afford.
//
// The durable requirement: the streaming writer must reach parity with the in-memory writer for
// image embedding — a registered image can be anchored onto a streamed worksheet, and the streamed
// package then carries the media part (real image bytes) and a drawing part anchoring it. This is a
// known-open gap today (the streamed worksheet has no addImage); the rewrite must close it.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-writer-image-parity',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'The streaming writer offers image parity with the in-memory writer: a registered image can be ' +
    'anchored onto a streamed worksheet, and the streamed package carries the media and drawing ' +
    'parts — so callers who need both out-of-core data and an embedded image are not forced in-core.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a streamed worksheet exposes an addImage anchor method (parity with the in-memory worksheet)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetAddImage} = await api.streamWriterImageSupport();
        assert.ok(
          sheetAddImage,
          'the streaming worksheet must expose addImage, like the in-memory worksheet',
        );
      },
    },
    {
      name: 'anchoring a registered image on a streamed sheet embeds the media and drawing parts',
      baseline: 'pass',
      async expect(api, assert) {
        const {error, mediaParts, drawingParts} = await api.streamWriterImageSupport('B2:D6');
        assert.strictEqual(
          error,
          null,
          `anchoring an image on a streamed sheet must not throw; got ${error}`,
        );
        assert.ok(mediaParts.length >= 1, 'the streamed package must carry the image media part');
        assert.ok(
          drawingParts.length >= 1,
          'the streamed package must carry a drawing part anchoring the image',
        );
      },
    },
  ],
};
