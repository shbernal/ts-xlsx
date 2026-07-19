// Cluster: styles
//
// Real-world scenario: a workbook has a conditional-formatting rule whose differential format
// (DXF) includes a custom number format. On a round-trip the DXF number format's code is
// serialized as the literal string "[object Object]" — the format-code object was coerced to a
// string instead of having its code read out — producing a styles part a strict consumer
// rejects. A DXF number format must round-trip as its real format code, never as "[object
// Object]".

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'conditional-format-numfmt-roundtrip/sample.xlsx';

export default {
  id: 'conditional-format-numfmt-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2698},
  cluster: 'styles',
  description:
    'A conditional-formatting differential format (DXF) that carries a custom number format ' +
    'round-trips with its real format code, never serialized as the literal string ' +
    '"[object Object]".',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'no DXF number format serializes as the literal "[object Object]"',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixtureStyleFacts(FIXTURE);
        assert.ok(source.dxfCount >= 1, 'precondition: the source has differential formats');
        const broken = rewritten.dxfFormatCodes.filter(c => /\[object Object\]/.test(c));
        assert.deepStrictEqual(broken, [], `no DXF numFmt may be "[object Object]"; got ${JSON.stringify(rewritten.dxfFormatCodes)}`);
      },
    },
  ],
};
