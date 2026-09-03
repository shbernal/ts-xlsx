// Cluster: streaming
//
// Real-world scenario: the streaming writer exists to hold a large sheet's memory flat, and it does
// that by serialising a row the moment it is committed and releasing its cells. Everything a cell
// carries that is serialised *inside* its `<row>` survives that. Everything serialised outside it did
// not, because the buffered pass gathers those by walking the sheet's rows at commit, and a committed
// row is no longer there to walk.
//
// A hyperlink's destination lives in the sheet's `<hyperlinks>` element plus an external
// relationship; a note lives in the comments and VML parts. Both were silently lost: the link's
// visible label came through as ordinary text with nothing to click, and the note's text vanished
// with no comments part at all. Nothing errored on either side.
//
// The third of the same shape is derived rather than carried: `collapsed="1"` rides a summary row
// only when its whole detail group is hidden, which is a question about the rows after it, so a
// streamed outline group rendered permanently expanded.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-streamed-row-keeps-what-lives-outside-it',
  provenance: {source: 'writer-parity-property'},
  cluster: 'streaming',
  description:
    'A row committed to the streaming writer keeps the content serialised outside its <row> ' +
    'element: its hyperlinks, its notes, and its place in an outline group whose collapsed state ' +
    'depends on rows other than itself. The package matches what the buffered writer produces for ' +
    'the same sheet.',

  behavior: [
    {
      name: 'a committed row emits its hyperlink, and the reader reads the destination back',
      async expect(api: CorpusApi, assert: Assert) {
        const {streamed} = await api.streamedRowSideContentReport();
        assert.deepEqual(streamed.hyperlinks, ['<hyperlink ref="A1" r:id="…" tooltip="go"/>']);
        assert.deepEqual(streamed.reread.a1, {
          text: 'link',
          hyperlink: 'https://example.com/a',
          tooltip: 'go',
        });
      },
    },
    {
      name: 'a committed row emits its note, so the comments and VML parts exist',
      async expect(api: CorpusApi, assert: Assert) {
        const {streamed} = await api.streamedRowSideContentReport();
        assert.ok(
          streamed.parts.some((name) => name.includes('comments')),
          `expected a comments part, got ${streamed.parts.join(', ')}`,
        );
        assert.ok(streamed.parts.some((name) => name.includes('vmlDrawing')));
        assert.equal(streamed.reread.b1Note, 'a note');
      },
    },
    {
      name: 'a committed summary row whose group is hidden still renders collapsed',
      async expect(api: CorpusApi, assert: Assert) {
        const {streamed, buffered} = await api.streamedRowSideContentReport();
        assert.deepEqual(streamed.rows, buffered.rows);
        assert.equal(streamed.sheetFormatPr, buffered.sheetFormatPr);
      },
    },
    {
      name: 'and the streamed package carries the same part families as the buffered one',
      async expect(api: CorpusApi, assert: Assert) {
        const {streamed, buffered} = await api.streamedRowSideContentReport();
        assert.deepEqual(streamed.parts, buffered.parts);
        assert.deepEqual(streamed.hyperlinks, buffered.hyperlinks);
      },
    },
  ],
} satisfies Case;
