// Cluster: tables
//
// Real-world scenario: a user sets an autofilter over a range. In Excel the filter dropdowns appear
// and work; in LibreOffice Calc the same file shows no filter at all. The reason is that a portable
// autofilter needs two coordinated declarations: the worksheet's <autoFilter> element AND a hidden
// workbook-level defined name `_xlnm._FilterDatabase`, scoped to the sheet, whose formula references
// the same range. Excel and portable writers emit both; a writer that only writes the worksheet
// autoFilter leaves LibreOffice unable to recognize the filter. Both must be present and consistent.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'autofilter-emits-filter-database-defined-name',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Setting an autofilter emits both the worksheet <autoFilter> ref and a hidden ' +
    '_xlnm._FilterDatabase defined name scoped to the sheet with a matching formula, so the filter is ' +
    'recognized by portable consumers (LibreOffice), not only Excel.',

  behavior: [
    {
      name: 'the worksheet autoFilter ref covers the requested range',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {autoFilterRef} = await api.autoFilterDefinedNameReport('A1:B2');
        assert.strictEqual(autoFilterRef, 'A1:B2', 'the autoFilter ref is written over the range');
      },
    },
    {
      name: 'a _xlnm._FilterDatabase defined name is declared for the autofilter',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasFilterDatabase} = await api.autoFilterDefinedNameReport('A1:B2');
        assert.strictEqual(
          hasFilterDatabase,
          true,
          'the _FilterDatabase defined name must be emitted so LibreOffice recognizes the filter',
        );
      },
    },
    {
      name: 'the _FilterDatabase defined name is hidden',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {filterDatabaseHidden} = await api.autoFilterDefinedNameReport('A1:B2');
        assert.strictEqual(
          filterDatabaseHidden,
          true,
          'the _FilterDatabase name is marked hidden, as Excel emits it',
        );
      },
    },
  ],
} satisfies Case;
