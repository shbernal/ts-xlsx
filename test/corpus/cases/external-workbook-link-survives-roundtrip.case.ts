// Cluster: xlsx-io
//
// Real-world scenario: a workbook links to a cell in another workbook — a formula or defined name
// resolves through `[1]Sheet!$A$1`, where `[1]` is an entry in the workbook's `<externalReferences>`
// backed by an `xl/externalLinks/externalLink1.xml` part whose own rels point (TargetMode="External")
// at the source file. The reader does not model external links, so before this fix a no-op load→save
// dropped the externalLink part and its `<externalReferences>` registration while KEEPING the formulas
// that use `[1]` — leaving a dangling external reference that Excel prompts to repair on open. The
// unmodeled external link (and its external-target pointer) must survive the round-trip intact.
//
// The fixture is synthetic and self-authored (no third-party bytes): a one-sheet package whose B1
// reads `[1]Questionnaire!$C$28`, with a minimal externalLink part linking a placeholder local path —
// distilled from a real 10-module .xlsm whose only round-trip repair prompt traced to this dropped link.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'external-workbook-link-survives-roundtrip/sample.xlsx';

export default {
  id: 'external-workbook-link-survives-roundtrip',
  provenance: {
    source: 'authored-repro',
    note: 'writer whole-package fidelity: dropped external link',
  },
  cluster: 'xlsx-io',
  description:
    'A no-op load→save preserves an external-workbook link — the externalLink part, its ' +
    'TargetMode="External" source pointer, and the <externalReferences> registration the [n] in a ' +
    'formula resolves through — rather than dropping it and dangling the reference.',

  behavior: [
    {
      name: 'the externalLink part and its <externalReferences> registration survive the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(
          source.externalLinks >= 1 && source.externalReferenceCount >= 1,
          'precondition: source has an external link and its <externalReferences> entry',
        );
        assert.strictEqual(
          rewritten.externalLinks,
          source.externalLinks,
          'externalLink parts survive',
        );
        assert.strictEqual(
          rewritten.externalReferenceCount,
          source.externalReferenceCount,
          '<externalReferences> registrations survive (no dangling [n])',
        );
      },
    },
    {
      name: 'the TargetMode="External" pointer to the source workbook survives verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(
          source.externalTargets.length >= 1,
          'precondition: source has an external target',
        );
        assert.deepStrictEqual(
          rewritten.externalTargets,
          source.externalTargets,
          'every external source-workbook pointer is re-emitted unchanged',
        );
      },
    },
  ],
} satisfies Case;
