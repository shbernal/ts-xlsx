// Cluster: data-validation
//
// Real-world scenario: a template builder applies a list-type dropdown to a whole column of cells so
// every row offers the same choices, sourcing the options from a range on a separate lookup sheet
// (e.g. Lookup!A1:A5). A recurring trap is the "shrinking dropdown": as the same validation is
// applied to more and more rows, the lower rows show progressively fewer options. The cause is
// relative-reference drift — if the source range is treated as relative to each target cell, the
// reference slides downward per row until it points past the end of the source data. The library must
// not do this: every targeted cell must persist the exact source-range reference it was given, so all
// cells expose the same complete option list regardless of how many rows are validated. (The writer
// is free to collapse the identical per-cell rules into a single sqref block — that is an efficiency
// win, not drift.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'list-validation-source-range-stable-across-rows',
  provenance: {source: 'upstream-issue'},
  cluster: 'data-validation',
  description:
    'Applying one list validation with a cross-sheet source range to a vertical span of cells ' +
    'persists the exact same source reference for every cell — no per-row relative drift — so the ' +
    'lowest row references the same full source range as the highest (dropdown options do not shrink ' +
    'with row index).',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every cell in the span keeps the exact cross-sheet source reference it was assigned',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, formulae, allIdentical} = await api.listValidationSourceRangeAcrossRows(
          6,
          'Lookup!A1:A5',
        );
        assert.strictEqual(
          allIdentical,
          true,
          `all rows must reference ${source} verbatim; got ${JSON.stringify(formulae)}`,
        );
      },
    },
    {
      name: 'the lowest row references the same source range as the highest (no drift past the source)',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, formulae} = await api.listValidationSourceRangeAcrossRows(6, 'Lookup!A1:A5');
        assert.strictEqual(formulae[0], source, 'first row source reference');
        assert.strictEqual(
          formulae[formulae.length - 1],
          source,
          'last row must not have drifted below the source',
        );
      },
    },
    {
      name: 'identical per-cell rules collapse into a single sqref block rather than one per row',
      baseline: 'pass',
      async expect(api, assert) {
        const {sqrefBlocks} = await api.listValidationSourceRangeAcrossRows(6, 'Lookup!A1:A5');
        assert.strictEqual(
          sqrefBlocks,
          1,
          'the six identical validations should serialize as one dataValidation with a spanning sqref',
        );
      },
    },
  ],
};
