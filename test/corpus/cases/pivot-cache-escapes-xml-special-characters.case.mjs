// Cluster: tables
//
// Real-world scenario: a pivot table is built over source data containing ordinary strings with
// XML-special characters — company names like "Smith & Co", angle brackets like "<West>", quoted
// text like It's "best" — plus rows with a missing (null) field value. The pivot cache serializes
// each distinct field value as a shared item. Those string values must be entity-escaped, or the
// pivotCacheDefinition part is malformed XML that Excel refuses to open. The bug: the writer emits
// the raw characters (a bare "&"), producing invalid XML. Missing values must serialize without
// throwing.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'pivot-cache-escapes-xml-special-characters',
  provenance: {source: 'upstream-pr'},
  cluster: 'tables',
  description:
    'A pivot cache built over source strings containing &, <, >, quotes serializes those values ' +
    'entity-escaped into well-formed pivotCacheDefinition XML, and tolerates null/missing source ' +
    'values, rather than emitting raw characters that corrupt the package.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a pivot over source data with special characters and a null value writes without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, writeError} = await api.pivotCacheSpecialCharsReport();
        assert.strictEqual(ok, true, `authoring the pivot must not throw; got ${JSON.stringify(writeError)}`);
      },
    },
    {
      name: 'the pivot cache XML is well-formed with special characters escaped',
      baseline: 'fail',
      async expect(api, assert) {
        const {cacheWellFormed, hasRawUnescapedAmp} = await api.pivotCacheSpecialCharsReport();
        assert.strictEqual(hasRawUnescapedAmp, false, 'no raw unescaped "&" leaks into the pivot cache XML');
        assert.strictEqual(cacheWellFormed, true, 'the pivotCacheDefinition part is well-formed XML');
      },
    },
  ],
};
