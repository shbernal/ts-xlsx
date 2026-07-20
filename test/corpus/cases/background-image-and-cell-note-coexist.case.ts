// Cluster: images
//
// Real-world scenario: a worksheet has both a sheet background image and at least one cell note
// (legacy comment). Both attach parts and relationships to the same worksheet — a note brings a
// comments part plus a legacy VML drawing, a background brings a picture part referenced from the
// worksheet relationships — so if their relationship wiring collides the written file is corrupt and
// strict applications refuse to open it. A worksheet declaring both must produce a package whose
// worksheet relationship ids are all unique and whose comment/VML and background parts are each
// present and independently referenced.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SPEC = {
  sheets: [
    {name: 'S', cells: [{ref: 'B2', value: 'x', note: 'a note'}], background: {extension: 'png'}},
  ],
};

export default {
  id: 'background-image-and-cell-note-coexist',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'A worksheet with both a background image and a cell note writes a valid package: the worksheet ' +
    'relationship ids are unique, the comments part and its VML drawing are present, and the ' +
    'background picture is referenced independently — no rel-id collision between the two features.',

  behavior: [
    {
      name: 'a worksheet with both a background image and a cell note writes without error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const result = await api.tryWriteWorkbook(SPEC);
        assert.strictEqual(
          result.ok,
          true,
          `writing must succeed; got ${JSON.stringify(result.error)}`,
        );
      },
    },
    {
      name: 'the worksheet relationship ids are all unique (no background/comment rel-id collision)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {consistency} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          consistency.worksheetRelIdsUnique,
          true,
          'no two worksheet relationships share an id',
        );
      },
    },
    {
      name: 'the comments part and its VML drawing are present alongside the background',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {packageParts, sheets} = await api.inspectPackage(SPEC);
        assert.ok(packageParts.hasCommentsPart, 'the note writes a comments part');
        assert.ok(packageParts.hasVmlDrawingPart, 'the note writes its VML drawing');
        assert.ok(sheets.S.hasBackgroundPicture, 'the worksheet references a background picture');
      },
    },
    {
      name: 'the background image is referenced by its own worksheet relationship',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {worksheetRels} = await api.inspectPackage(SPEC);
        assert.ok(
          worksheetRels.some((r: CorpusApi) => r.type === 'image'),
          'an image-type worksheet relationship backs the background picture',
        );
      },
    },
  ],
} satisfies Case;
