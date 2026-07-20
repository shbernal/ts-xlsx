---
name: validate-ooxml
description: Validate generated .xlsx output against Microsoft's OOXML schema/semantic oracle (OpenXmlValidator). Use after changing any writer path, when a workbook "opens broken in Excel", when checking that emitted XML conforms to ECMA-376, or when you need an authoritative pass/fail on a generated package. Pairs with write-corpus-case; see docs/agent-correctness-playbook.md.
---

# Validating OOXML output

The corpus (`pnpm run corpus`) proves well-formedness, package structure, and no
regression — but it is *our* assertions. The **authoritative** schema + semantic check
is Microsoft's own `OpenXmlValidator`, wrapped in a repo-owned .NET tool (ADR-0002). It
is the independent oracle: if it says a package is clean, Excel's own conformance layer
agrees. Reach for it after any change to a writer path.

## Requirement

It needs **.NET 10** (`dotnet` on PATH). The published npm package has zero .NET
dependency — this is development-only tooling. If you don't have .NET, do **not**
substitute an XSD/`xmllint` validator (deliberately not wired — `schemas/README.md`,
ADR-0002). Fall back to `pnpm run corpus` locally and let CI's `ooxml-validation`
workflow run the oracle on your PR. See `docs/agent-correctness-playbook.md`.

## Validate a file you already have

```bash
pnpm run validate:ooxml -- path/to/workbook.xlsx another.xlsx
```

Deterministic JSON to stdout. Exit codes: **0** every input clean · **1** validation or
package-open errors found · **2** the tool could not run (bad args / internal failure).
Each error carries `id`, `type` (`Schema` | `Semantic` | `Package`), `partUri`, and
`xpath` — enough to locate the offending element without opening the zip.

## Emit a representative workbook, then validate

Node 24 runs the `.ts` sources directly, so a repro imports straight from `src`. Write
one that exercises the feature you changed, then point the validator at it:

```js
// /tmp/repro.mjs
import {writeFileSync} from 'node:fs';
import {Workbook} from './src/core/workbook.ts';
import {writeXlsx} from './src/io/xlsx/write.ts';

const wb = new Workbook();
const ws = wb.addWorksheet('Data');
ws.addRow(['Name', 'Value']);
ws.addRow(['alpha', 42]);
// …exercise the exact path you touched: styles, tables, formulas, images, CF…

writeFileSync('/tmp/repro.xlsx', writeXlsx(wb));   // writeXlsx → Uint8Array
```

```bash
node /tmp/repro.mjs && pnpm run validate:ooxml -- /tmp/repro.xlsx
```

For the streaming writer, use `WorkbookStreamWriter` (`src/io/xlsx/write-stream.ts`) and
collect its output the same way, then validate — streaming and buffered output must both
be clean.

## The baseline (do not paper over new errors)

`test/ooxml-validation/run.ts` (`pnpm run test:ooxml`) generates buffered + streaming
workbooks and control cases, then compares each package's errors to the frozen set in
`test/ooxml-validation/allowed-errors.json`.

- A **new** diagnostic that isn't baselined → the gate fails. Fix the writer; do not add
  it to the baseline to silence it. The baseline is a record of *known-open* writer bugs
  we've chosen to track, not a mute button.
- A **stale** baseline (an error you *fixed*) also fails the gate — remove that entry in
  the same change so the file stays honest.
- Match errors by the exact `{id, type, partUri, xpath}` fingerprint.

## After validating

- Clean (exit 0)? Good — but the oracle is not a substitute for the corpus. If the change
  is new behavior, also land an implementation-blind corpus case (`write-corpus-case`
  skill) so the guarantee is permanent and survives without .NET.
- Errors (exit 1)? The `xpath` + `partUri` point at the element. Fix the serializer, not
  the assertion. Re-run until clean.
