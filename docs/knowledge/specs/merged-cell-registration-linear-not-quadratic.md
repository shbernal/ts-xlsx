# Merged-cell registration must scale linearly, not quadratically

Cluster: merged-cells / performance

## Scenario

A worksheet contains tens of thousands of merged ranges (real files reach 30k+ — for example one
small horizontal merge per row across thousands of rows). Loading such a file takes minutes, or
effectively never finishes in constrained/browser environments, even though the data is trivial.
The correctness of merge preservation is covered by the `many-merged-cells-preserved-and-overlap-
rejected` corpus case; this note is about the parse *time*.

> Captured as a spec note, not a corpus case: exercising the real 30k-merge fixture would make the
> corpus run take minutes. A bounded-time check belongs in a perf harness with a hard timeout, not
> the behavior corpus.

## Root cause

The naive loader registers merges one at a time and, for each new merge, scans **all**
previously-registered merges to reject overlaps — an O(merges²) collision check. With tens of
thousands of merges this dominates load time.

## Desired behavior

- Registering merged ranges during load scales roughly linearly (or n·log n) with the number of
  merges, not quadratically.
- A merge only ever conflicts with a cell already part of another merge, so the collision check
  should consult a **per-cell "already merged" index** (each cell records its master), making each
  registration proportional to its own area rather than to the count of prior merges.

## Open questions

- Is overlap validation even required when parsing a well-formed foreign file? Excel guarantees
  non-overlapping merges in valid OOXML, so load could trust the input and only validate on
  user-driven `mergeCells` calls, keeping a spatial index for the mutation API.
- Should a load-time overlap conflict throw, or be tolerated/repaired, given the fork's
  foreign-generator-tolerance stance? (Consistency with `namespace-agnostic-bom-tolerant-ooxml-
  reading` and the other tolerance notes.)
- What is the target: a memory-bounded per-cell index vs. an interval tree — and which wins for
  the common "one merge per row" shape.

Related: `whole-column-data-validation-bounded-memory` (the same "don't materialize per-cell state
for a range feature" principle).
