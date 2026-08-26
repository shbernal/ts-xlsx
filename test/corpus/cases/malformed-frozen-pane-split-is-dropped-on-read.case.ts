// Cluster: sheet-views
//
// Real-world scenario: a `<pane>` element spells its split as a fraction, a negative, or a word --
// the shapes a hand-edited file or a foreign generator produces. `Worksheet.freeze()` refuses all
// three as non-negative integers, so a reader that stores the attribute verbatim leaves the model
// holding a value its own authoring API would not accept, and the failure surfaces much later out
// of the serializer as a RangeError naming a column the file never mentioned. A malformed split
// must be dropped at the point it is read, leaving the rest of the sheet view intact.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'malformed-frozen-pane-split-is-dropped-on-read',
  provenance: {source: 'writer-boundary-probe'},
  cluster: 'sheet-views',
  description:
    'A <pane> split attribute that is not a non-negative integer is dropped on read rather than ' +
    'stored, so the loaded model never holds a split the authoring API refuses and the workbook ' +
    'it was read from writes back out without throwing.',

  behavior: [
    {
      name: 'a malformed split never reaches the model',
      expect(api: CorpusApi, assert: Assert) {
        for (const {spelling, xSplit, ySplit} of api.malformedFrozenSplitReport()) {
          for (const [axis, split] of [
            ['xSplit', xSplit],
            ['ySplit', ySplit],
          ] as const) {
            assert.ok(
              split === null || (Number.isInteger(split) && split >= 0),
              `xSplit="${spelling}" left ${axis} as ${String(split)}, which freeze() would refuse`,
            );
          }
        }
      },
    },
    {
      name: 'and the workbook it was read from writes back out',
      expect(api: CorpusApi, assert: Assert) {
        for (const {spelling, rewriteError} of api.malformedFrozenSplitReport()) {
          assert.equal(
            rewriteError,
            null,
            `a package read with xSplit="${spelling}" must re-write; got ${String(rewriteError)}`,
          );
        }
      },
    },
  ],
} satisfies Case;
