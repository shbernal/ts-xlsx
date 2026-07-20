// Cluster: formulas
//
// Real-world scenario: a workbook is built through the streaming writer with a master formula cell
// (e.g. B1 = A1*2) and a block of shared-formula slave cells (B2..B10 each referencing B1 as their
// shared-formula master) — the standard compact way to apply one formula down a column. When the
// streamed package is read back, the slave cells must be real formula cells (resolving to the shared
// formula, or at least a value), not empty. A streaming writer that drops the slaves leaves every
// derived cell blank in the output.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streaming-shared-formula-slaves-populate',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A streaming-written worksheet with a master formula and shared-formula slave cells reloads with ' +
    'the master intact and the slave cells populated (a resolved formula/value), not emptied.',

  behavior: [
    {
      name: 'the master formula cell survives the streaming write',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {masterHasFormula} = await api.streamingSharedFormulaReport();
        assert.strictEqual(
          masterHasFormula,
          true,
          'the master formula cell reloads as a formula cell',
        );
      },
    },
    {
      name: 'shared-formula slave cells are not dropped to empty on a streaming write',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {slaveResolved, slaveValue} = await api.streamingSharedFormulaReport();
        assert.strictEqual(
          slaveResolved,
          true,
          `a shared-formula slave must reload populated, not empty; got ${JSON.stringify(slaveValue)}`,
        );
      },
    },
  ],
} satisfies Case;
