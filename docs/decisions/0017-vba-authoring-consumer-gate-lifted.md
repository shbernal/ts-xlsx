# ADR 0017 — VBA authoring is in scope; the consumer gate is lifted

**Status:** Accepted (2026-07-23) · supersedes the "deferred pending a forcing consumer" clause of
ADR 0016 (and the parallel gate in ADR 0014) for the VBA feature line.

## Context

ADRs 0013/0014/0016 share one governing move: an authoring surface stays *deferred* until a concrete
**forcing consumer** exists — a caller whose real use pins down the API shape, the invariants, and the
edge cases. The reasoning (CLAUDE.md §4, no premature abstraction) is sound: you cannot validate the
ergonomics of a write path in a vacuum, and a guessed-at surface is a guaranteed future break. For VBA
specifically, ADR 0016 shipped a **read-only** view and listed a value-to-cost forward map for
authoring — attach-blob, then first-class source→bytes — each item gated on that consumer.

The project owner has now removed the gate for the VBA line: they are the consumer, the directive is to
build the feature out to its natural completion ("max features"), and where a slice needs a real file
to prove itself, we author a fixture. This does **not** repeal the no-premature-abstraction principle —
it satisfies its underlying requirement by a different, and arguably stronger, means.

## Decision

1. **A crafted fixture under the strict test suite is an accepted substitute for a forcing consumer.**
   The gate exists to guarantee an authoring surface is *validated against real use before it sets*. A
   fixture that must survive read→write→re-read byte-faithfully — and, where feasible, pass Excel's own
   open/repair check or the OOXML validator (ADR 0012/0013 tiers) — exercises the surface at least as
   hard as a casual caller would. Feature work no longer waits on an external consumer appearing.

2. **Authoring lands as vertical slices, each fully green, never one speculative drop.** In the
   ADR-0016 cost order:
   - **§2.1 attach-blob (this slice, done):** `Workbook.vbaProjectBytes` — a get/set accessor pair over
     the raw `vbaProject.bin`. The getter returns a defensive copy of the attached blob (or `undefined`
     for a macro-free workbook); the setter attaches/replaces it, or removes the project when set to
     `undefined`. A set is **validated fail-closed** (`parseVbaProject` must accept the bytes *before*
     any state changes) so a malformed blob is rejected with `VbaParseError` and never half-applied.
     Replacing or removing drops the previous blob's whole closure, so a now-stale `vbaProjectSignature`
     over the old bytes is discarded rather than left advertising a broken signature. The attached
     reference is byte-identical in shape to what the reader captures for a macro workbook, so the
     writer emits a valid macro-enabled package with **zero writer changes** — the content-type and
     rel-wiring machinery already keys off the preserved `vbaProject` reference.
   - **§2.3 first-class authoring (in progress):** synthesize a valid `vbaProject.bin` from
     edited/created module source. This makes `VbaProject` (or a sibling authoring surface) an **emission
     authority**, which reverses ADR 0016's "read-only view, not a source of truth" core; it therefore
     gets its own ADR amending 0016 when the *authoring surface* lands (§2.3d), not a silent extension of
     this one. It is built as internal encode primitives first — each provable in isolation before any
     public shape changes:
     - **§2.3a CFB writer (`writeCompoundFile`, done):** the encode counterpart to `cfb.ts` — a
       hierarchy of storages and streams → a v3 [MS-CFB] container. Emits each storage's children as the
       name-ordered balanced tree a host *navigates* (not just the linear scan our own reader uses), so
       the modules resolve under the `VBA` storage in Excel. Proven by re-encoding a real 156 KB
       Excel-authored project stream-identical, by an independent directory-tree walk reaching every
       entry, and by `parseVbaProject` decoding the result. Internal to `src/vba`; not on the public
       barrel yet.
     - **§2.3b MS-OVBA compressor (`compressContainer`, done):** the encode inverse of
       `decompressContainer` — the copy-token/literal run compression Excel emits, with a raw-chunk
       fallback when a window would not shrink. Proven by round-trip at every chunk boundary, by
       run-length overlap collapsing repetitive data, and by re-expanding to the exact bytes of a real
       Excel-compressed `dir` stream (it even edges out Excel's own output on that stream). Internal.
     - **§2.3c `dir`/`PROJECT`/module synthesis, §2.3d Workbook authoring surface:** the remaining
       assembly layer and the public surface, each a further green slice. §2.3d is the authority shift
       and carries the ADR-0016 amendment.

3. **The read/attach path stays the safety floor.** §2.1 leaves ADR 0016's read invariant intact:
   preservation is still the sole emission authority, `vbaProjectBytes` simply lets a caller *supply*
   those preserved bytes instead of only receiving them from a read. Nothing about the read view or the
   byte-faithful passthrough regresses.

4. **Executing VBA remains permanently out of scope** (ADR 0013): running macros needs a live host and
   is never a document-library feature. "Authoring" here means *producing valid bytes*, not *running
   them*.

## Consequences

- **Positive:** a caller can now attach, replace, copy, or strip a macro project through one honest
  accessor pair — `dst.vbaProjectBytes = src.vbaProjectBytes` copies macros between workbooks; setting
  `undefined` demotes an `.xlsm` to a plain package; an externally-produced `.bin` imports in one line.
  All validated fail-closed, all fixture-backed.
- **Scope discipline preserved, not abandoned.** The gate's *intent* (don't set an unvalidated surface)
  still binds — it is now met by fixtures + the strict suite rather than by waiting. Slices that would
  bound coverage still say so; nothing ships un-green.
- **The big reversal is still ahead and still gets its own ADR.** §2.3 making an authoring surface the
  emission authority is the genuine shape-change to ADR 0016; this ADR only opens the door and lands the
  thin, non-authority-shifting first slice.
- **Revisit when:** §2.3 lands (amend ADR 0016), or if a real consumer's needs contradict a shape a
  fixture led us to — in which case the fixture was too weak a proxy and the surface breaks, cheaply,
  under SemVer.

Related: ADR 0016 (VBA read view), ADR 0014 (round-trip-only charts/shapes/slicers — same gate, not yet
lifted), ADR 0013 (Excel as a test oracle, never a runtime), `docs/knowledge/specs/xlsm-macro-preservation.md`.
