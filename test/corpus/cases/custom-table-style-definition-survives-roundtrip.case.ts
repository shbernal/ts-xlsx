// Cluster: styles
//
// Real-world scenario: a workbook defines its own table style — a `<tableStyle>` in styles.xml whose
// elements paint the whole table, the header row, the stripes — and a table asks for it by name in
// `tableStyleInfo/@name`. A no-op round-trip that regenerates styles.xml without the `<tableStyles>`
// block leaves the table asking for a style nothing defines: it opens without complaint and renders
// completely unstyled, so the file looks "fine" while every brand colour is gone.
//
// The definitions are preserved verbatim, which makes two things load-bearing:
//
//  • A `tableStyleElement`'s `dxfId` is an *index* into the differential-style table. That table is
//    re-emitted at its original indices, which is the only reason a verbatim fragment's references
//    stay meaningful — renumbering it would silently re-point every element at a different format.
//    So the case asserts on the dxf each element actually lands on, not on the index.
//  • A verbatim fragment carries its namespace prefixes with it. Excel stamps `xr9:uid` on every
//    table style it writes, so a stylesheet root that declares only the default namespace makes the
//    part unparseable — a far louder failure than the dropped style this preservation prevents.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'custom-table-style-definition-survives-roundtrip/branded-table-style.xlsx';

export default {
  id: 'custom-table-style-definition-survives-roundtrip',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'A workbook’s custom `<tableStyle>` definitions, and the default table/pivot styles it ' +
    'nominates, survive a no-op round-trip — so a table referencing a custom style by name still ' +
    'resolves to a real definition instead of rendering unstyled.',

  behavior: [
    {
      name: 'a table’s custom style name still resolves to a definition after a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.strictEqual(
          source.nameOnTable,
          'Harbour Table',
          'precondition: the table asks for a custom style by name',
        );
        assert.strictEqual(
          source.tableStyleOnTableResolves,
          true,
          'precondition: the source defines that style',
        );
        assert.strictEqual(
          rewritten.nameOnTable,
          source.nameOnTable,
          'the table still asks for it',
        );
        assert.strictEqual(
          rewritten.tableStyleOnTableResolves,
          true,
          'and the re-emitted stylesheet still defines it',
        );
      },
    },
    {
      name: 'each style element still points at the same differential style',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.deepStrictEqual(
          rewritten.elements,
          source.elements,
          'every tableStyleElement survives with its type and dxfId',
        );
        // The load-bearing half: the dxf indices must still mean what they meant. Comparing the
        // resolved fragments — not the indices — is what makes a renumbered dxf table fail here.
        assert.deepStrictEqual(
          rewritten.elementDxfs,
          source.elementDxfs,
          'and each dxfId still resolves to the same differential style',
        );
        assert.ok(
          source.elementDxfs.every((dxf: CorpusApi) => dxf !== null),
          'precondition: every element’s dxfId resolves in the source',
        );
      },
    },
    {
      name: 'the nominated default table and pivot styles survive',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.strictEqual(rewritten.defaultTableStyle, source.defaultTableStyle);
        assert.strictEqual(rewritten.defaultPivotStyle, source.defaultPivotStyle);
      },
    },
    {
      name: 'the re-emitted stylesheet declares every prefix its preserved fragments use',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.deepStrictEqual(
          source.undeclaredPrefixes,
          [],
          'precondition: the source part is namespace-well-formed',
        );
        assert.deepStrictEqual(
          rewritten.undeclaredPrefixes,
          [],
          'a preserved fragment’s prefix (Excel’s xr9:uid) must not be left undeclared',
        );
      },
    },
  ],
} satisfies Case;
