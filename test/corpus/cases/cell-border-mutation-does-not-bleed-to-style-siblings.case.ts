// Cluster: styles
//
// Real-world scenario: a workbook (often a loaded template) has many cells that reference the same
// shared style/xf record because they were formatted identically. The user assigns a border to one
// cell, expecting only that cell to gain a border. Because the cells alias one style object, mutating
// one cell's border mutates the shared record and the border appears on every sibling that shared it.
// Correct behavior is copy-on-write: a per-cell border assignment isolates that cell's style at the
// point of mutation. (This is the border facet of the shared-style aliasing family — see
// per-cell-fill-isolation and shared-base-style-font-mutation-isolated.)

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'cell-border-mutation-does-not-bleed-to-style-siblings',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Assigning a border to one cell that shares a style record with siblings borders only that cell — ' +
    'the siblings keep their unbordered style, with no shared-record bleed.',

  behavior: [
    {
      name: 'the targeted cell gains the border',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {a1} = await api.loadMutateCellBorder();
        assert.strictEqual(a1, true, 'the cell the border was assigned to has it');
      },
    },
    {
      name: 'style-sharing siblings do not gain an unrequested border',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {a2, a3, bled} = await api.loadMutateCellBorder();
        assert.strictEqual(bled, false, `no sibling may gain the border; got A2=${a2} A3=${a3}`);
      },
    },
  ],
} satisfies Case;
