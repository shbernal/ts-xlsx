# Carrying image-bearing content between workbooks must copy the media or fail loudly

Cluster: images

## Scenario

A user copies content — a worksheet, a range, a drawing — from one workbook into another, and the
copied images render as broken in the destination file. The cause is structural: an image in OOXML is
a two-part thing — a drawing/anchor in the worksheet that references an image by a workbook-scoped
media id, plus the actual media bytes and the relationship that binds them. Copying only the
drawing/anchor (or the cell model) into a different workbook carries a reference to an image id that
does not exist there, so the writer emits a package with a dangling drawing relationship pointing at
media that was never registered — a silently broken image.

> Spec note, not a corpus case: the report is a support question with no reproduction, code, or
> fixture; the one durable nugget is the diagnostic that images are a media+relationship pair, not a
> value that travels with a copied cell/drawing. The durable value is the cross-workbook copy
> contract; a corpus case follows once a copy affordance exists to exercise.

## Desired behavior

- **Copying image-bearing content transparently copies the media.** When a drawing/anchor (or a
  worksheet/range containing one) is copied into a different workbook, the underlying image media
  bytes and their relationships are registered in the destination workbook and the anchor is rebound
  to the destination's media id — so the image renders, not breaks.
- **Or fail loudly, never silently.** If the library cannot or does not copy the media (e.g. a
  low-level model copy that only moves the anchor), an anchor referencing a media id absent from the
  destination workbook must produce a clear, actionable error at write time — naming the missing image
  — rather than emitting a package with a dangling drawing relationship that a consumer shows as a
  broken image.
- **Round-trip integrity of the destination package.** After a cross-workbook copy, every drawing
  relationship in the written package resolves to a real media part with a unique rel id (the same
  packageParts/rel-id invariant the table/comment coexistence cases assert), and re-reading the
  destination surfaces the image anchored where it was placed.

## Open questions

- What is the public copy affordance this hangs off — a worksheet/range copy API, a `model`-level
  assignment, or a dedicated `copyImagesTo(destWorkbook)`? The contract differs: a high-level copy
  should carry media automatically; a raw `model` splice cannot and should error.
- De-duplication: if the same image is copied into a destination that already holds identical media,
  is it registered once (shared) or duplicated? Prefer content-hash de-dup to avoid bloating the
  package.
- Scope: does "content copy" include floating drawings only, or also header/footer images, background
  images, and (later) in-cell rich-value images — each of which has its own media wiring?

Related: `add-image-source-contract`, `image-embedded-in-cell-vs-floating-anchor`,
`streaming-write-add-image`, `comment-and-table-coexist-on-same-sheet`.
