// Cluster: styles
//
// Real-world scenario: the alignment enums are a typed contract — VerticalAlignment is
// top|center|bottom|justify|distributed, and TypeScript rejects anything else at the call site.
// But an untyped JavaScript caller (or any value smuggled past the types) can hand the writer an
// out-of-contract token such as vertical="middle" — the ExcelJS-era spelling of what OOXML calls
// "center". Three oracles disagree on the identical bytes that produces: real Excel opens the file
// and silently ignores the attribute; the OOXML schema rejects it (ST_VerticalAlignment has no
// "middle"); and the library's own reader — which narrows every alignment token through a guard —
// silently drops it on read-back. The writer trusting the types and passing the value straight
// through is therefore a silent asymmetry: it emits a schema-invalid file whose own reader discards
// the styling, with no error anywhere. The durable contract: a value the writer accepts must survive
// read-back — an out-of-contract enum value must be rejected at the write boundary or preserved,
// never silently serialized into a file the library itself cannot read back.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'out-of-contract-alignment-enum-not-silently-dropped',
  cluster: 'styles',
  description:
    'An in-contract vertical alignment value round-trips; an out-of-contract enum value (one the ' +
    'types forbid but an untyped caller can still pass) must not be silently written into a ' +
    "schema-invalid file the library's own reader then discards — the writer either rejects it at " +
    'its boundary or preserves it, but never emits-then-drops it without an error.',
  provenance: {source: 'writer-boundary-probe'},

  behavior: [
    {
      name: 'an in-contract vertical="center" round-trips through write and read',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeThrew, readBackVertical} = await api.alignmentVerticalEnumReport('center');
        assert.strictEqual(
          writeThrew,
          false,
          'a valid vertical alignment must write without error',
        );
        assert.strictEqual(
          readBackVertical,
          'center',
          'a valid vertical alignment must survive the round-trip',
        );
      },
    },
    {
      name: 'an out-of-contract vertical="middle" is not silently written-then-dropped',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeThrew, readBackVertical} = await api.alignmentVerticalEnumReport('middle');
        // Correct behavior is either boundary: reject the value at write time (writeThrew), or carry
        // it through so it reads back (preserved, or coerced to the canonical "center"). The one
        // outcome the contract forbids is a clean write whose value vanishes on read — that is the
        // silent asymmetry this case locks out.
        assert.ok(
          writeThrew || readBackVertical === 'middle' || readBackVertical === 'center',
          `an out-of-contract vertical value must be rejected or preserved, never silently dropped; ` +
            `got writeThrew=${writeThrew}, readBackVertical=${JSON.stringify(readBackVertical)}`,
        );
      },
    },
  ],
} satisfies Case;
