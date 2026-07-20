// Cluster: styles
//
// Real-world scenario: a loaded workbook deduplicates identical formatting, so cells saved with the
// same style share one in-memory style object. Setting a single style facet on one cell via its
// setter must change only that cell — but if the setter mutates the shared object in place, every
// sibling that happened to share the style silently inherits the change. The fill, font, and border
// facets of this aliasing family are locked separately (per-cell-fill-isolation,
// shared-base-style-font-mutation-isolated, cell-border-mutation-does-not-bleed-to-style-siblings);
// this case covers the remaining facets — alignment, number format, and protection — for which the
// setter today edits the aliased record and bleeds into the sibling. Correct behavior is
// copy-on-write: the assignment isolates the target cell's style at the point of mutation.

import type {Assert, Behavior, Case, CorpusApi} from '../case.ts';

const FACETS = ['alignment', 'numFmt', 'protection'];

const isolation = FACETS.flatMap((facet): Behavior[] => [
  {
    name: `setting ${facet} on one cell changes that cell (control)`,
    baseline: 'pass',
    async expect(api: CorpusApi, assert: Assert) {
      const {target, original} = await api.loadMutateCellFacet(facet);
      assert.notStrictEqual(target, original, `the edited cell must reflect the new ${facet}`);
    },
  },
  {
    name: `setting ${facet} on one cell does not bleed into a style-sharing sibling (in memory)`,
    baseline: 'pass',
    async expect(api: CorpusApi, assert: Assert) {
      const {sibling, original, bled} = await api.loadMutateCellFacet(facet);
      assert.ok(
        !bled,
        `the sibling must keep its original ${facet}; it changed from ${original} to ${sibling}`,
      );
    },
  },
  {
    name: `after write-back only the edited cell's ${facet} changed on disk`,
    baseline: 'pass',
    async expect(api: CorpusApi, assert: Assert) {
      const {diskSibling, diskBled} = await api.loadMutateCellFacet(facet);
      assert.ok(
        !diskBled,
        `the sibling must keep its original ${facet} in the written file; it became ${diskSibling}`,
      );
    },
  },
]);

export default {
  id: 'cell-style-setter-isolates-alignment-numfmt-protection',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Setting the alignment, number format, or protection of one loaded cell that shares a style ' +
    'record with siblings changes only that cell — the setter is copy-on-write and never mutates ' +
    'the aliased shared style, in memory or after write-back.',

  behavior: isolation,
} satisfies Case;
