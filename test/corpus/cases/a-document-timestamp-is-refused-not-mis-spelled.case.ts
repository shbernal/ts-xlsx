// Cluster: dates
//
// Real-world scenario: a caller sets `workbook.properties.created` from a parsed string, a database
// column, or arithmetic on another date. Two of the ways that goes wrong produce a `Date` that is
// still a `Date`: an Invalid Date (truthy, `instanceof Date`, passes every guard), and a year outside
// 0000-9999. `corePropsXml` called `toISOString()` on whatever arrived, so the first threw a bare
// `RangeError: Invalid time value` -- outside this library's taxonomy, naming neither the property
// nor the document -- and the second silently wrote ISO 8601's expanded `+275760-09-13T…` form,
// which is not in the lexical space of `dcterms:W3CDTF` and makes Excel repair the package.
//
// The rule this locks: a document timestamp that cannot be spelled is refused as an `AuthoringError`
// naming the property, the way every other unwritable value on this path already is, and never
// written as something the schema does not admit.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-document-timestamp-is-refused-not-mis-spelled',
  provenance: {source: 'write-path-audit'},
  cluster: 'dates',
  description:
    'An Invalid Date or an out-of-range year in a core document property raises an AuthoringError ' +
    'naming the property rather than a bare RangeError or an expanded-year timestamp no W3CDTF ' +
    'can carry; an ordinary one is written to the second, as Excel writes it.',

  behavior: [
    {
      name: 'an ordinary timestamp is written to the second, in both properties',
      expect(api: CorpusApi, assert: Assert) {
        const rows = api.documentDateReport().filter((row) => row.kind === 'an ordinary timestamp');
        assert.equal(rows.length, 2, 'created and modified');
        for (const row of rows) {
          assert.equal(row.threw, false, `${row.property}: writes`);
          assert.equal(
            row.stamp,
            '2026-01-02T03:04:05Z',
            `${row.property}: no millisecond field, which is what Excel emits`,
          );
        }
      },
    },
    {
      name: 'a Date with no spelling is refused, inside the taxonomy, naming the property',
      expect(api: CorpusApi, assert: Assert) {
        const rows = api.documentDateReport().filter((row) => row.kind !== 'an ordinary timestamp');
        assert.equal(rows.length, 6, 'three broken shapes across two properties');
        for (const row of rows) {
          assert.equal(row.threw, true, `${row.property} / ${row.kind}: refused`);
          assert.equal(
            row.isXlsxError,
            true,
            `${row.property} / ${row.kind}: catchable as XlsxError, not a bare ${row.errorName}`,
          );
          assert.equal(row.errorName, 'AuthoringError', `${row.property} / ${row.kind}: authoring`);
          assert.ok(
            row.message?.includes(row.property),
            `${row.property} / ${row.kind}: the message names the property (${row.message})`,
          );
          assert.equal(row.stamp, null, `${row.property} / ${row.kind}: nothing was written`);
        }
      },
    },
  ],
} satisfies Case;
