# ADR 0012 — Three tiers of correctness evidence; round-trip proves consistency, not conformance

**Status:** Accepted (2026-07-21) · Phase 4 · sharpens [ADR 0002](./0002-ooxml-validation-oracle.md) (validation oracle) and [ADR 0010](./0010-agent-correctness-dispatch.md) (correctness dispatch)

## Context

Most of the corpus writes a workbook with our writer and reads it back with our
reader — `roundtripWorkbook` and the bulk of the bespoke capabilities. ADR 0002
already noted the hazard in one line ("a writer and reader can agree on malformed
OOXML") and stood up the `OpenXmlValidator` oracle in response. The table-header bug
(`9579f152`) showed the hazard is sharper than that line implies, and that even the
oracle does not close it.

A table declared its column names only in `xl/tables/table1.xml` and left the header
row missing from the worksheet grid. Our reader takes column names from the table
part, so **write→read was a perfect fixed point** — every round-trip case stayed
green. The structural facts (`inspectPackage`) were green too, because they were
partitioned by package part: each part was internally valid. And the schema oracle
was green, because each part conforms to ECMA-376 in isolation — the rule Excel
enforces (the header cells must exist and carry the column names) is a *cross-part
behavioral* rule the written spec does not state. The only thing that rejected the
file was Excel Desktop.

The lesson is not "add one more fact." It is that a green corpus was systematically
misread as evidence of conformance when it was mostly evidence of self-consistency,
and that the gap between the two is where "files that do not open in Excel" lives.

## Decision

Treat correctness evidence as **three tiers of strictly increasing authority and
cost**, and never let a lower tier stand in for a claim only a higher tier can make.

### Tier 1 — Self-consistency (`R(W(x)) ≈ x`)

Write with our writer, read with our reader, assert the model survives. This is a
**fixed-point property of our own code.** It proves the encoder and decoder share a
model; it proves nothing about whether that model matches Excel.

- **Catches:** *unilateral* bugs — one half malformed and the other half strict about
  it (a dropped field, an unparseable emission).
- **Structurally blind to:** *correlated* bugs, where writer and reader are wrong in
  compensating directions or ignore the same field. These are not a rare tail: writer
  and reader for a feature are typically authored together from one mental model, so
  their blind spots are correlated *by construction*. The table bug is the canonical
  instance.
- **Cheapest, highest volume, weakest.** Sufficient for **intra-model** invariants (a
  value survives, a style does not bleed) and nothing more.

### Tier 2 — Spec-conformance (independent oracle)

Microsoft's `OpenXmlValidator` (ADR 0002), plus the structural `inspectPackage` facts.
An **independent implementation** breaks the write/read correlation — it is not our
code, so it cannot share our blind spot.

- **Catches:** everything the written standard (ECMA-376) states — schema validity,
  OPC part structure, enums, child-element ordering, per-part semantic rules.
- **Ceiling:** it enforces what the spec *says*, not what Excel *does*. Excel's real
  acceptance rules are a **superset** of the spec, and many are cross-part invariants
  the spec never spells out. The table bug passed the oracle.
- Required for any **single-part conformance** claim; round-trip is not enough there.

### Tier 3 — Excel behavior (the application itself)

What Excel Desktop actually does with the file. The **only** ground truth for Excel's
quirks and for cross-part invariants the spec omits. Not cheaply automatable; reached
by hand and captured as provenance (`{source: 'excel-desktop-verification'}`).

- The table invariant was **discovered** here and could only be discovered here.
- Required for any **cross-part correspondence** or known-Excel-quirk claim.

### The rule

**Match the tier to the invariant, and record which tier witnessed it.**

| Invariant class | Lowest sufficient tier | Example |
| --- | --- | --- |
| Intra-model (value/style survives a round-trip) | Tier 1 | a number keeps its type; a font does not bleed |
| Single-part conformance (well-formed, ordered, valid enum) | Tier 2 | `CT_Worksheet` child order; valid `patternType` |
| Cross-part correspondence / Excel quirk | Tier 3 (seed) + a Tier-2 seam fact (lock) | table columns ↔ header cells |

A cross-part invariant is **seeded** by one Tier-3 verification (prove Excel enforces
it) and then **locked** by a corpus fact whose *shape is the relationship itself*, so
regressions are caught in CI without re-opening Excel. The header fix did exactly this:
one Excel-Desktop verification, then two new seam facts (`cellText`, `columnNames`)
that phrase the correspondence.

### Corollary: the corpus is partitioned by part, and that is a blind spot

`inspectPackage` describes each package part well and the relationships *between* parts
barely. Any invariant spanning two parts is unstateable until a fact is deliberately
built to cross the seam. An audit of the current vocabulary (2026-07-21) found these
**unstated cross-part seams** Excel is known or expected to enforce:

- ~~`<mergeCell ref>` ↔ covered cells must be empty (no `mergeCells` fact).~~
  **Closed** — `mergeCleanReport` fact + `merged-range-opens-without-repair-prompt.case.ts`.
- `<tablePart r:id>` in a sheet → rel → table part → content-type override (tables carry
  no owning sheet or rel id).
- Drawing/image anchor → `r:embed` → drawing rel → media part → content-type (no
  drawing-XML parsing or media enumeration).
- `<dataValidation sqref>` ↔ the cells it claims (no `dataValidation` fact).
- Hyperlink ↔ external rel vs internal location (no hyperlink fact).
- `<sheet r:id>` → rel → existing worksheet part (captured but only checked part→declaration).
- Cell `s=` index → `styles.xml` xf/font/fill (`cellText` drops `s`; only the default font
  is inspected).
- ~~Shared-formula master ↔ slaves (`formulas` holds `<f>` text only, not `t="shared" si= ref=`).~~
  **Closed** — `sharedFormulas` fact + `shared-formula-master-slave-geometry-structural.case.ts`
  lock the master/slave `si`/`ref` geometry structurally. **Left open (see hazard below):** the
  `ref` for a *non-contiguous* clone set.
- Comment ↔ VML shape ↔ `legacyDrawing` rel.

**Structural ceiling:** `worksheetRels` reads only `sheet1.xml.rels`, and tables/drawings
are not tied to an owning sheet. Every *multi-sheet* cross-part chain and every
"reference-in-sheet → rel → target part" resolution is unstateable until that is lifted —
it is the precondition for several rows above, so it is the highest-leverage fix.

These are not scheduled here; they are the backlog this ADR makes visible. Close each
the same way: Tier-3 seed, then a seam fact that locks it.

### Open Tier-3 hazard: shared-formula `ref` over-covers a non-contiguous clone set

Discovered while closing the shared-formula seam (2026-07-21). `planSharedFormulas`
(`src/io/xlsx/worksheet-xml.ts`) computes a master's `ref` as the **bounding rectangle**
of the master plus all its clones. For the common fills — down a column, across a row —
that rectangle *is* the clone set, and the emitted geometry is exact. But when the clones
are non-contiguous, the rectangle covers cells that were never cloned:

```
master B1, clones B2 + D5  →  <f t="shared" ref="B1:D5" si="0">…</f>
```

`B1:D5` is fifteen cells; only three (`B1`, `B2`, `D5`) carry a `<c>`. The other twelve are
absent from the sheet. The hazard is that the shared-formula `ref` is, to some consumers, an
*instruction to materialize the formula across the whole rectangle*: **LibreOffice** auto-fills
every cell in `ref` with the translated formula, so those twelve empty cells would silently
gain a formula the caller never wrote. This is the same shape as the merge-repair bug — the
writer emits geometry a consuming app interprets more aggressively than we intended — which is
why it is called out rather than left implicit.

It is deliberately **not locked and not fixed**:

- **Not locked.** A seam fact asserting current behavior would freeze a `ref` we suspect is
  wrong. Per the recipe, a cross-part invariant is locked only *after* a Tier-3 seed proves what
  the correct geometry is; locking first would cement the bug.
- **Not fixed blind.** The correct behavior is unknown without Tier 3, and the fix is not cheap:
  a rectangular `ref` cannot represent a non-contiguous group, so a correct writer must either
  split the group into maximal contiguous runs (each its own master, each needing the formula
  **translated** to its new anchor — the R1C1/relative-offset machinery we do not yet have) or
  degrade non-fill clones to standalone `<f>` cells. Both are real work; neither should be
  attempted before the Tier-3 seed says which is even needed.

**What a Tier-3 seed requires — how to confirm Excel's exact behavior.** The tier is irreducible:
there is no local oracle for it (the schema validator passed the analogous table bug). It means,
concretely:

1. **Emit the probe file.** Write a `.xlsx` whose only shared group is non-contiguous — e.g.
   master `B1` with clones `B2` and `D5`, so ts-xlsx produces `ref="B1:D5"` with twelve empty
   interior cells. (`scratch-totals-probe.ts` is the kind of throwaway harness this uses.)
2. **Open it in Excel Desktop** — the real application, not Online or a viewer, since the quirk
   is an application acceptance rule, not a spec rule — and observe, for the twelve empty cells,
   which of these Excel actually does:
   - **Materializes** them (auto-fills the translated formula across the rectangle, like
     LibreOffice) → our `ref` is actively wrong; it injects data. Fix is mandatory.
   - **Ignores** the interior and keeps only the three written cells → the over-wide `ref` is
     cosmetically loose but harmless in Excel; the fix is then driven by other consumers
     (LibreOffice) rather than Excel.
   - **Repairs** the file (the "we found a problem… recover?" dialog) → the geometry is rejected
     outright, the highest-severity outcome, same class as the merge bug.
3. **Cross-check the other direction:** re-save from Excel and diff the `ref` Excel itself writes
   for the same non-contiguous group — that reveals the geometry Excel considers canonical (most
   likely: it never emits a non-contiguous group as one shared formula at all).
4. **Record the observed behavior** as `provenance: {source: 'excel-desktop-verification'}` on the
   seeding case, then lock it with a seam fact whose shape *is* the corrected `ref`↔clone-set
   relationship, so CI catches regressions without re-opening Excel.

Until someone with Excel Desktop runs steps 1–3, the correct geometry is unknown, so the seam is
recorded here rather than closed.

## Consequences

- **Positive:** "the corpus is green" is no longer mistaken for "Excel accepts this."
  The playbook's defense-in-depth table now carries the epistemic distinction, not just
  a cost ordering. New cross-part invariants have a named recipe (seed + seam fact) and a
  visible backlog.
- **Neutral:** cross-part correctness still costs a manual Excel verification to seed.
  That cost is real and intended — Tier 3 is irreducible; the seam fact is what keeps it
  a one-time cost per invariant.
- **Does not change** the single-oracle stance (ADR 0010): Tier 2 remains exactly one
  independent oracle. This ADR adds no validator; it names what each existing tier can and
  cannot witness.
