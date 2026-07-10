# Duplicate a contiguous block of rows, N copies at a time

Cluster: tables

## Scenario

A caller wants to duplicate not a single row but a contiguous run of rows as a group — e.g. copy a
5-row template block and paste it 10 times to build a repeating layout — in one operation, rather
than looping a single-row duplicate. Today the duplication affordance handles one row (and a repeat
count for that one row); there is no first-class way to copy a multi-row block as a unit and stamp it
several times at a chosen destination.

> Spec note, not a corpus case: this is a feature-proposal generalization of an existing single-row
> operation, not a bug with a reproduction. The durable value is the operation's shape and the
> semantics it must preserve. When implemented it composes the existing row-op vocabulary and becomes
> assertable by reading back the pasted range's values/formulas/styles/heights.

## Desired behavior

- **A block-duplicate operation copies a contiguous run of rows as a group.** Parameters (names
  illustrative): the first source row of the block, the block length (number of rows in the group),
  the destination row where the first copy is pasted, and the number of copies to generate. Copies
  are pasted contiguously in sequence starting at the destination.
- **Single-row duplication remains a special case** (length 1, one copy) and behaves exactly as the
  existing single-row duplicate does — the block form generalizes it without breaking it.
- **Each copy is faithful.** Cell values, formulas, number formats, fonts/fills/borders, row heights,
  and outline levels of the source block are reproduced in every copy. Merged regions wholly inside
  the block are re-created per copy; the merge/ref bookkeeping stays consistent (the same discipline
  the row-insert/duplicate cases already exercise for merges and table/image refs).
- **Formula adjustment follows the existing rule.** Relative references in copied formulas shift by
  the paste offset per copy, matching the single-row duplicate's relative-reference behavior; absolute
  references stay fixed. (See `duplicate-row-formula-relative-adjust`.)
- **Interaction with structured content is defined.** Duplicating a block that overlaps or extends a
  table, or that shifts anchored images/defined names below it, must keep the package valid and the
  refs consistent — or reject with a clear error if the operation would corrupt a table.

## Open questions

- Insert-and-shift vs. overwrite: does pasting copies push existing rows down (insert semantics) or
  overwrite the destination rows? Insert semantics matches "duplicate rows" intent; make it explicit.
- Paste location relative to the source: is pasting a block onto a destination that overlaps the
  source range allowed, and if so what are the semantics?
- Does the copy include row-level data validations, conditional formatting ranges, and comments
  anchored in the block, or only cell content and style?

Related: `duplicate-row-formula-relative-adjust`, `row-insert-and-duplicate-shift-merged-cells`,
`duplicate-row-copies-faithfully-and-permits-merge`, `streaming-writer-worksheet-splice-rows-columns`.
