// Cluster: formulas
//
// A shared formula is stored once on a master cell as `<f t="shared" ref="B1:B3" si="0">A1*2</f>`;
// every other cell in the group is a slave carrying only `<f t="shared" si="0"/>` — no formula text,
// just a back-reference to the master's index. Excel enforces the correspondence structurally: a
// slave's `si` must resolve to a master that declared it, distinct groups must carry distinct `si`,
// and a slave must sit inside its master's `ref` range. A dangling `si`, a duplicated `si`, or a
// slave outside the ref is the geometry Excel repairs on open.
//
// The rest of the shared-formula corpus asserts this through a write→read round-trip, which is a
// fixed point of our own encoder/decoder pair: the reader resolves a slave by its `si`, so a writer
// that stamped the wrong `si`/`ref` and a reader that read it back the same wrong way would agree and
// the round-trip would stay green (ADR 0012, tier 1). This case reads the geometry straight off the
// emitted `<f>` elements — an independent structural witness of the master↔slave relationship that a
// correlated writer/reader bug cannot hide from.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Two independent groups on one sheet: a fill-down (B1 master over B2/B3) and a second, distinct
// group (C1 master over C2). Two groups are what makes `si` uniqueness observable.
const TWO_GROUPS = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A2', value: 2},
        {ref: 'A3', value: 3},
        {ref: 'B1', formula: 'A1*2', result: 2},
        {ref: 'B2', sharedFormula: 'B1', result: 4},
        {ref: 'B3', sharedFormula: 'B1', result: 6},
        {ref: 'C1', formula: 'A1*3', result: 3},
        {ref: 'C2', sharedFormula: 'C1', result: 6},
      ],
    },
  ],
};

export default {
  id: 'shared-formula-master-slave-geometry-structural',
  provenance: {source: 'cross-part-seam-audit', ref: 'ADR-0012'},
  cluster: 'formulas',
  description:
    'The shared-formula master/slave geometry emitted to the worksheet — a master carrying ' +
    '`t="shared" ref si` with its formula text, and slaves carrying only a matching `si` — is ' +
    'asserted structurally, not through the reader round-trip, so a correlated writer/reader `si`/' +
    '`ref` bug cannot pass unseen.',

  behavior: [
    {
      name: 'each group has a master carrying its ref range and formula text',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {masters} = (await api.inspectPackage(TWO_GROUPS)).sheets.S.sharedFormulas;
        const byCell: Record<string, CorpusApi> = Object.fromEntries(
          masters.map((m: CorpusApi) => [m.cell, m]),
        );
        assert.strictEqual(
          masters.length,
          2,
          `two masters expected; got ${JSON.stringify(masters)}`,
        );
        assert.strictEqual(byCell.B1?.ref, 'B1:B3', 'the fill-down master spans B1:B3');
        assert.strictEqual(byCell.C1?.ref, 'C1:C2', 'the second master spans C1:C2');
      },
    },
    {
      name: 'distinct groups carry distinct shared indices',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {masters, siUniqueAcrossMasters} = (await api.inspectPackage(TWO_GROUPS)).sheets.S
          .sharedFormulas;
        assert.strictEqual(
          siUniqueAcrossMasters,
          true,
          `two masters must not share a si; got ${JSON.stringify(masters)}`,
        );
      },
    },
    {
      name: 'every slave references a real master and sits inside its ref',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {slaves, everySlaveHasMaster, everySlaveWithinMasterRef} = (
          await api.inspectPackage(TWO_GROUPS)
        ).sheets.S.sharedFormulas;
        assert.strictEqual(
          slaves.length,
          3,
          `three slaves expected; got ${JSON.stringify(slaves)}`,
        );
        assert.strictEqual(
          everySlaveHasMaster,
          true,
          `every slave si must resolve to a master; got ${JSON.stringify(slaves)}`,
        );
        assert.strictEqual(
          everySlaveWithinMasterRef,
          true,
          `every slave must fall inside its master's ref; got ${JSON.stringify(slaves)}`,
        );
      },
    },
  ],
} satisfies Case;
