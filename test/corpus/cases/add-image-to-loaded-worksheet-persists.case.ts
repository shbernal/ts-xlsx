// Cluster: images
//
// Real-world scenario: a user opens an existing file (parsed from bytes into an in-memory workbook),
// registers a new image on the workbook, and anchors it onto an existing worksheet of the loaded
// file. On re-serialization the added image must actually be present — the media part, a drawing part
// describing the anchor, and the worksheet-to-drawing relationship — and re-reading must surface it.
// The same add-image flow must behave identically whether the target worksheet came from a loaded
// package or was created fresh, guarding against loaded worksheets silently dropping new drawings.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'add-image-to-loaded-worksheet-persists',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'Registering an image on a workbook loaded from bytes and anchoring it onto an existing loaded ' +
    'worksheet persists the image on re-serialization — the output carries the media and drawing ' +
    'parts and re-reads as one image.',

  behavior: [
    {
      name: 'the re-serialized package includes the added image media',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasMedia} = await api.addImageToLoadedWorksheetReport();
        assert.strictEqual(hasMedia, true, 'the media part for the added image is present');
      },
    },
    {
      name: 'the re-serialized package includes the drawing part describing the anchor',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasDrawing} = await api.addImageToLoadedWorksheetReport();
        assert.strictEqual(hasDrawing, true, 'the drawing part linking the anchor is present');
      },
    },
    {
      name: 're-reading the output surfaces exactly the one added image',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadImageCount} = await api.addImageToLoadedWorksheetReport();
        assert.strictEqual(reloadImageCount, 1, 'the added image is enumerated after reload');
      },
    },
  ],
} satisfies Case;
