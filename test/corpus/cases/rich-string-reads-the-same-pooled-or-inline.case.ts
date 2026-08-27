// Cluster: xlsx-io
//
// Real-world scenario: whether a rich string arrives pooled in `xl/sharedStrings.xml` or inline in
// the worksheet body is a producer's encoding choice, not a property of the content. Excel pools;
// a streaming writer often does not; the same document saved by two tools differs on it. A reader
// that answers differently for the two hands the caller a different cell value for the same cell,
// decided by something the author never chose.
//
// The two are separate grammars in the format, and were separate hand-written copies of one grammar
// in this reader, kept in step by a comment saying they agreed. Nothing measured it: each reader is
// exercised through its own path, so the drift this locks against would have shipped silently.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'rich-string-reads-the-same-pooled-or-inline',
  provenance: {source: 'hand-authored'},
  cluster: 'xlsx-io',
  description:
    'One rich string, read once from a pooled <si> and once from an inline <is>, produces an ' +
    'identical cell value: per-run text, per-run fonts and run order alike.',

  behavior: [
    {
      name: 'a pooled and an inline rich string read to the same value',
      expect(api: CorpusApi, assert: Assert) {
        const {pooled, inline, identical} = api.richStringPooledVersusInlineReport();
        assert.equal(identical, true, `pooled ${pooled} !== inline ${inline}`);
      },
    },
    {
      name: 'and that value carries every run with its own formatting',
      expect(api: CorpusApi, assert: Assert) {
        // Pinning the reading itself, not only that the two agree: two readers agreeing on the
        // wrong answer would satisfy the check above on its own.
        const {pooled} = api.richStringPooledVersusInlineReport();
        assert.deepEqual(JSON.parse(pooled), {
          richText: [
            {
              font: {bold: true, color: {argb: 'FFFF0000'}, name: 'Arial', size: 14},
              text: 'bold red',
            },
            {text: ' plain '},
            {font: {italic: true, underline: true}, text: 'italic'},
          ],
        });
      },
    },
  ],
} satisfies Case;
