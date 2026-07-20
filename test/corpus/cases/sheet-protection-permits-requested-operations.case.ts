// Cluster: styles
//
// Real-world scenario: an author protects a worksheet but explicitly opts to leave some operations
// available to end users — sorting, autofilter, and formatting cells/rows/columns — so the structure
// is locked while people can still sort and filter the data. OOXML's <sheetProtection> uses INVERTED
// booleans: an attribute value of "1" LOCKS an operation and "0" (or omission) PERMITS it. So a
// caller asking to keep sorting available must produce sort="0", not sort="1". If the permissive
// options were dropped or inverted, the protected sheet would forbid exactly the operations the
// author meant to allow.

import type {Assert, Case, CorpusApi} from '../case.ts';

const PERMISSIVE = {
  password: 'pw',
  options: {sort: true, autoFilter: true, formatCells: true, formatColumns: true, formatRows: true},
};

export default {
  id: 'sheet-protection-permits-requested-operations',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Protecting a worksheet while opting to keep sort, autofilter, and cell/row/column formatting ' +
    'available emits a <sheetProtection> with the sheet locked but those operations permitted ' +
    '(attribute "0", the OOXML "not forbidden" encoding) — the permissive options are honored, not ' +
    'dropped or inverted to a lock.',

  behavior: [
    {
      name: 'protecting the sheet emits a sheetProtection element with protection enabled',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetProtectionAttrs} = await api.authorCellProtection(
          [{ref: 'A1', value: 'x'}],
          PERMISSIVE,
        );
        assert.ok(sheetProtectionAttrs, 'a <sheetProtection> element is written');
        assert.strictEqual(
          sheetProtectionAttrs.sheet,
          '1',
          'the sheet itself is protected (sheet="1")',
        );
      },
    },
    {
      name: 'sorting and autofilter are permitted (not forbidden) under protection',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetProtectionAttrs} = await api.authorCellProtection(
          [{ref: 'A1', value: 'x'}],
          PERMISSIVE,
        );
        assert.strictEqual(
          sheetProtectionAttrs.sort,
          '0',
          'sort is permitted (sort="0"), not locked',
        );
        assert.strictEqual(
          sheetProtectionAttrs.autoFilter,
          '0',
          'autofilter is permitted (autoFilter="0")',
        );
      },
    },
    {
      name: 'formatting of cells, rows, and columns is permitted under protection',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetProtectionAttrs} = await api.authorCellProtection(
          [{ref: 'A1', value: 'x'}],
          PERMISSIVE,
        );
        assert.strictEqual(sheetProtectionAttrs.formatCells, '0', 'formatting cells is permitted');
        assert.strictEqual(
          sheetProtectionAttrs.formatColumns,
          '0',
          'formatting columns is permitted',
        );
        assert.strictEqual(sheetProtectionAttrs.formatRows, '0', 'formatting rows is permitted');
      },
    },
  ],
} satisfies Case;
