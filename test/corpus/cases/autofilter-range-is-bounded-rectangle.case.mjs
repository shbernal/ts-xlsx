// Cluster: tables
//
// Real-world scenario: a user enables an auto filter over a data range and later sorts by it.
// Sorting only works if the written autoFilter reference is a bounded rectangle that covers the
// data — both start and end row and column present. A column-only, row-unbounded reference (like
// "A:AZ") makes sorting fail. When a bounded range is applied, the library must emit that exact
// bounded reference and round-trip it unchanged, and it must stay within the sheet dimension.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'H1'},
        {ref: 'B1', value: 'H2'},
        {ref: 'C1', value: 'H3'},
        {ref: 'A2', value: 1},
        {ref: 'B2', value: 2},
        {ref: 'C2', value: 3},
        {ref: 'A3', value: 4},
        {ref: 'B3', value: 5},
        {ref: 'C3', value: 6},
      ],
      autoFilter: 'A1:C3',
    },
  ],
};

const isBoundedRect = (ref) => /^[A-Z]+\d+:[A-Z]+\d+$/.test(ref || '');

export default {
  id: 'autofilter-range-is-bounded-rectangle',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2289},
  cluster: 'tables',
  description:
    'An auto filter applied over a bounded data range is emitted as that exact bounded ' +
    'rectangle (both start and end row and column) and round-trips unchanged — never a ' +
    'column-only, row-unbounded reference that breaks sorting.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the emitted autoFilter ref is the applied bounded rectangle',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          sheets.S.autoFilterRef,
          'A1:C3',
          'the exact bounded range is serialized',
        );
        assert.ok(isBoundedRect(sheets.S.autoFilterRef), 'the ref has both row and column bounds');
      },
    },
    {
      name: 'the autoFilter range round-trips unchanged',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.autoFilter,
          'A1:C3',
          'the bounded filter range survives read→write',
        );
      },
    },
  ],
};
