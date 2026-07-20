// Cluster: types
//
// Real-world scenario: a cell is assigned a rich-text value made of several formatted runs, one of
// which carries an empty text string (a middle run whose text is ""). An empty run serializes to a
// run element with an empty `<t>` payload, which Excel's schema rejects — opening the file prompts a
// corrupt-file recovery. An empty run contributes nothing to the rendered text, so dropping it at
// serialization time produces a valid file whose visible appearance is unchanged. The writer must
// omit zero-length runs rather than emit an empty `<t>` element.

import type {Assert, Case, CorpusApi} from '../case.ts';

const RUNS = [
  {text: 'a', font: {bold: true}},
  {text: '', font: {italic: true}},
  {text: 'b', font: {}},
];

export default {
  id: 'rich-text-empty-substring-run',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A rich-text value containing a run with an empty text string is written without emitting a ' +
    'zero-length run element (which Excel flags as corrupt); the empty run is dropped while the ' +
    'surrounding non-empty runs keep their text and formatting.',

  behavior: [
    {
      name: 'no empty <t> run element is serialized for a run with empty text',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {emptyTextRunInXml} = await api.richTextRoundtripReport(RUNS);
        assert.strictEqual(
          emptyTextRunInXml,
          false,
          'an empty-text run must not be emitted as an empty <t> element',
        );
      },
    },
    {
      name: 'the empty-text run is dropped, leaving only the non-empty runs',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {runs} = await api.richTextRoundtripReport(RUNS);
        assert.deepStrictEqual(
          runs.map((r: CorpusApi) => r.text),
          ['a', 'b'],
          'only the two non-empty runs survive',
        );
      },
    },
    {
      name: 'the non-empty runs surrounding the dropped run keep their text and formatting (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {runs} = await api.richTextRoundtripReport(RUNS);
        const a = runs.find((r: CorpusApi) => r.text === 'a');
        const b = runs.find((r: CorpusApi) => r.text === 'b');
        assert.ok(a && a.bold === true, 'the leading bold run survives with its formatting');
        assert.ok(b, 'the trailing run survives');
      },
    },
  ],
} satisfies Case;
