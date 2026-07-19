// Cluster: xlsx-io
//
// Real-world scenario: a worksheet's page header and footer are configured with distinct content
// for different page classes — separate text for the first page, for even-numbered pages, and for
// odd pages. In OOXML the `<headerFooter>` element carries `firstHeader`/`firstFooter`,
// `evenHeader`/`evenFooter`, and `oddHeader`/`oddFooter` children, but the first- and even-page
// variants are only honored by consuming applications when the gating attributes are set:
// `differentFirst="1"` for first-page content and `differentOddEven="1"` for even-page content.
// Without those flags a spreadsheet application shows the odd content on every page even though the
// even/first elements are present — the community workaround (duplicating odd content into even)
// exists precisely because the differentiation flags are missing.
//
// The invariant: emitting the variant child elements is not enough; the gating flags must be set
// whenever the corresponding variant is provided, or the variants are silently ignored downstream.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [{ref: 'A1', value: 'x'}],
      headerFooter: {
        oddHeader: 'ODD-H', oddFooter: 'ODD-F',
        evenHeader: 'EVEN-H', evenFooter: 'EVEN-F',
        firstHeader: 'FIRST-H', firstFooter: 'FIRST-F',
      },
    },
  ],
};

export default {
  id: 'headerfooter-first-even-variants-written',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'A worksheet configured with first-, even-, and odd-page header/footer variants writes each ' +
    'variant child element AND sets the gating differentFirst / differentOddEven attributes, ' +
    'without which consuming applications ignore the first- and even-page content.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the odd, even, and first header/footer child elements are all emitted with their text',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        const hf = sheets.S.headerFooter;
        assert.strictEqual(hf.oddHeader, 'ODD-H', 'odd header text is written');
        assert.strictEqual(hf.evenHeader, 'EVEN-H', 'even header text is written');
        assert.strictEqual(hf.firstHeader, 'FIRST-H', 'first header text is written');
        assert.strictEqual(hf.firstFooter, 'FIRST-F', 'first footer text is written');
      },
    },
    {
      name: 'differentOddEven is set when even-page header/footer content is provided',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          sheets.S.headerFooter.differentOddEven,
          true,
          'even variants require differentOddEven="1" or they are ignored on open'
        );
      },
    },
    {
      name: 'differentFirst is set when first-page header/footer content is provided',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          sheets.S.headerFooter.differentFirst,
          true,
          'first-page variants require differentFirst="1" or they are ignored on open'
        );
      },
    },
  ],
};
