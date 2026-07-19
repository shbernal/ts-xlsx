// Cluster: xlsx-io
//
// Real-world scenario: some spreadsheet generators write the document core-properties part
// (docProps/core.xml) using the core-properties namespace as the DEFAULT xml namespace
// (xmlns=".../core-properties") rather than binding it to the conventional "cp" prefix. Elements
// like lastModifiedBy and lastPrinted then appear UNPREFIXED (<lastModifiedBy>) instead of as
// <cp:lastModifiedBy>, while created/creator keep their own dc/dcterms prefixes. A reader that
// matches core-property elements by literal prefixed tag name treats the unprefixed spelling as an
// unexpected node and throws while parsing — so an otherwise valid workbook fails to open at all.
// Core properties must be matched by namespace + local name, so the prefixed and default-namespace
// spellings are read identically.
//
// The fixture is a real workbook whose docProps/core.xml declares the core-properties namespace as
// the default xmlns and emits an unprefixed lastModifiedBy / lastPrinted (the library's own writer
// always emits the cp: prefix, so the foreign spelling is injected).

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'core-properties-unprefixed-namespace-elements/foreign-core.xlsx';

export default {
  id: 'core-properties-unprefixed-namespace-elements',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'A workbook whose docProps/core.xml declares the core-properties namespace as the default xmlns ' +
    '(so lastModifiedBy / lastPrinted are unprefixed) loads without throwing, and its lastModifiedBy ' +
    'value is read — core properties are matched by namespace and local name, not literal prefix.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook with an unprefixed, default-namespace core.xml loads without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(ok, true, `the workbook must load; instead the reader threw: ${error}`);
      },
    },
    {
      name: 'the lastModifiedBy value from the unprefixed core.xml is read',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, lastModifiedBy} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(ok, true, 'precondition: the workbook loads');
        assert.strictEqual(lastModifiedBy, 'Editor Two', `lastModifiedBy must be read from the unprefixed element; got ${JSON.stringify(lastModifiedBy)}`);
      },
    },
  ],
};
