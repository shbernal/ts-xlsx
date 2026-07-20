// Cluster: styles
//
// Real-world scenario: a cell holds rich text made of multiple runs where the very first run carries
// character formatting (e.g. an underlined leading word) and the rest is plain. On read-back the
// leading run's formatting must survive — a run's position within the cell text must not decide
// whether its formatting is preserved. This locks that formatting on a leading run is read back
// identically to the same formatting on a later run.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'richtext-leading-run-formatting-preserved',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Character formatting on the leading run of a rich-text cell survives a round-trip, identically ' +
    'to the same formatting on a non-leading run — run position does not affect formatting survival.',

  behavior: [
    {
      name: 'a formatted leading run keeps its underline after round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {runs} = await api.richTextRoundtripReport([
          {text: 'here', font: {underline: true}},
          {text: ' plain', font: {}},
        ]);
        const lead = runs.find((r: CorpusApi) => r.text === 'here');
        assert.ok(
          lead && lead.underline === true,
          'the underlined leading run survives with its underline',
        );
      },
    },
    {
      name: 'the same formatting on a non-leading run is preserved identically (position independence)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {runs} = await api.richTextRoundtripReport([
          {text: 'plain ', font: {}},
          {text: 'here', font: {underline: true}},
        ]);
        const tail = runs.find((r: CorpusApi) => r.text === 'here');
        assert.ok(
          tail && tail.underline === true,
          'the underlined non-leading run survives with its underline',
        );
      },
    },
    {
      name: 'the leading and trailing runs are read back as distinct runs',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {runCount} = await api.richTextRoundtripReport([
          {text: 'here', font: {underline: true}},
          {text: ' plain', font: {}},
        ]);
        assert.strictEqual(runCount, 2, 'both runs are surfaced distinctly');
      },
    },
  ],
} satisfies Case;
