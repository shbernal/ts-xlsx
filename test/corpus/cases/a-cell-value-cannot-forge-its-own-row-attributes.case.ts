// Cluster: streaming
//
// Real-world scenario: an export streams a grouped report, and one of the summary rows carries a
// cell whose text happens to be ` collapsed="1"` -- a snippet pasted out of a spec, a log line, a
// column of XML fragments. `collapsed="1"` is the one row attribute that cannot be decided when the
// row is serialised, because it is a fact about the rows *after* it, so a streamed summary row is
// patched once its detail group is known. The patcher used to ask the rendered `<row>` markup whether
// the attribute was already there -- and that markup also holds the row's cell text, which
// `escapeText` leaves double quotes in verbatim. So the cell answered the question for the row, and
// the outline group rendered expanded with nothing anywhere reporting a problem.
//
// The rule this locks: content never decides how its own container is serialised. What the row
// declared is carried as a field, not re-read out of the bytes it was rendered into.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-cell-value-cannot-forge-its-own-row-attributes',
  provenance: {source: 'write-path-audit'},
  cluster: 'streaming',
  description:
    'A streamed collapsed summary row is serialised identically whatever its cells contain: a cell ' +
    'whose value spells a row attribute cannot suppress that attribute.',

  behavior: [
    {
      name: 'a collapsed group is marked collapsed on the summary row',
      async expect(api: CorpusApi, assert: Assert) {
        const {control} = await api.collapsedSummaryUnderPoisonedCellText();
        assert.ok(
          control.includes(' collapsed="1"'),
          `the control summary row carries the attribute (got ${control})`,
        );
      },
    },
    {
      name: 'a cell value spelling the attribute does not change the row that holds it',
      async expect(api: CorpusApi, assert: Assert) {
        const {poisoned, control} = await api.collapsedSummaryUnderPoisonedCellText();
        assert.equal(
          poisoned,
          control,
          'the row opens identically whatever the cell beneath it says',
        );
      },
    },
  ],
} satisfies Case;
