// Cluster: rows
//
// Real-world scenario: an author builds an outlined (grouped) set of rows and wants the group
// collapsed when the file first opens. In OOXML a correctly collapsed outline needs two coordinated
// facts: every detail (child) row carries its outlineLevel and is hidden, AND the summary row that
// terminates the group carries the `collapsed` flag on its own row element. When the collapsed flag
// is instead written on the hidden detail rows (and omitted from the summary row), the rows show as
// hidden but the expand/collapse toggle is out of sync — the user must click twice to expand. The
// detail rows get their outlineLevel + hidden right today; the collapsed flag is placed on the wrong
// rows.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Detail rows 2–4 grouped under summary row 5 (summary-below, the default). Rows 2–4 are hidden
// detail; row 5 is the summary that should carry the collapse toggle.
const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'top'},
        {ref: 'A2', value: 'd2'},
        {ref: 'A3', value: 'd3'},
        {ref: 'A4', value: 'd4'},
        {ref: 'A5', value: 'summary'},
      ],
      rows: [
        {index: 2, outlineLevel: 1, hidden: true},
        {index: 3, outlineLevel: 1, hidden: true},
        {index: 4, outlineLevel: 1, hidden: true},
      ],
    },
  ],
};

export default {
  id: 'row-outline-collapsed-flag-belongs-on-summary-row',
  provenance: {source: 'upstream-issue'},
  cluster: 'rows',
  description:
    'A collapsed row-outline group writes its detail rows with outlineLevel + hidden (correct today) ' +
    'and places the collapsed toggle on the summary row that terminates the group — not on the ' +
    'hidden detail rows, where it is emitted today, leaving the expand/collapse control out of sync.',

  behavior: [
    {
      name: 'each detail row carries its outline level and is hidden',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        for (const r of ['2', '3', '4']) {
          assert.strictEqual(
            sheets.S.rows[r].outlineLevel,
            1,
            `detail row ${r} has outlineLevel 1`,
          );
          assert.strictEqual(sheets.S.rows[r].hidden, true, `detail row ${r} is hidden`);
        }
      },
    },
    {
      name: 'the hidden detail rows do NOT carry the collapsed flag',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        for (const r of ['2', '3', '4']) {
          assert.strictEqual(
            sheets.S.rows[r].collapsed,
            false,
            `the collapsed toggle belongs on the summary row, not on hidden detail row ${r}`,
          );
        }
      },
    },
    {
      name: 'the summary row terminating the group carries the collapsed flag',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          sheets.S.rows['5'] ? sheets.S.rows['5'].collapsed : false,
          true,
          'the summary row (row 5) must carry collapsed="1" so the outline toggle expands in one click',
        );
      },
    },
  ],
} satisfies Case;
