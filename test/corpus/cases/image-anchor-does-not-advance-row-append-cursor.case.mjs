export default {
  id: 'image-anchor-does-not-advance-row-append-cursor',
  cluster: 'images',
  description:
    'Anchoring a floating image over a cell range is a drawing overlay, not a row insertion. It must ' +
    'not advance the worksheet row-append cursor: a caller who anchors an image at A1:B3 and then ' +
    'appends rows expects those rows to fill the sheet from the top, exactly as if no image were ' +
    'present, and the final layout must not depend on whether the image or the rows were added first. ' +
    'A regression made anchoring an image push subsequent appended rows below the anchored range, ' +
    'leaving the top rows empty.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'appending rows before anchoring an image places data at the top (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {rowsFirst} = await api.imageAnchorRowAppendReport();
        assert.strictEqual(rowsFirst.firstDataCell, 'a');
      },
    },
    {
      name: 'anchoring an image first still places subsequently appended rows at the top',
      baseline: 'pass',
      async expect(api, assert) {
        const {imageFirst} = await api.imageAnchorRowAppendReport();
        assert.strictEqual(imageFirst.firstDataCell, 'a');
      },
    },
    {
      name: 'the appended-row layout is identical regardless of image-vs-rows add order',
      baseline: 'pass',
      async expect(api, assert) {
        const {imageFirst, rowsFirst} = await api.imageAnchorRowAppendReport();
        assert.strictEqual(imageFirst.firstDataCell, rowsFirst.firstDataCell);
      },
    },
  ],
};
