# Embedding images in the streaming write path

Cluster: streaming

## Scenario

A user generates very large workbooks with the streaming (row-by-row) write path to keep memory
bounded, and needs to embed images — for example a logo anchored over a cell range on each sheet. The
buffered document-write path serializes images as drawing parts with anchors, but the streaming path
has no equivalent, so adding an image fails outright. Users are forced to choose between memory-safe
streaming and image support, which are mutually exclusive today.

> Spec note, not a corpus case: the feature does not exist yet, so there is no behavior to assert —
> but the adapter surface is ready to lock it the moment it lands (see below). The durable value is
> the design reconciling image placement with the streaming lifecycle.

## Desired behavior

- The streaming writer supports embedding images with the **same anchoring model** as the buffered
  path: one-cell, two-cell, and absolute anchors; range / `tl`+`br` and `tl`+`ext` forms.
- The design reconciles placement with the streaming lifecycle: image **media** (binary + content
  type / part) can be registered at any time, but the **drawing relationship and `xdr` anchor XML**
  for a sheet must be emitted when that sheet is committed. So image additions for a sheet must be
  accepted **before its commit point** — queued, then flushed with the sheet's drawing part and
  worksheet rel.
- The resulting package is **byte-structurally equivalent** to the buffered path's: a drawing part
  per sheet, a worksheet→drawing relationship, image media parts under the media folder, unique
  relationship ids, and correct content-type overrides/defaults for the image formats.

## Prior art / test-readiness

The buffered path already implements worksheet image addition and anchor serialization, and a
third-party fork added `addImage` to the streaming worksheet, so the feature is feasible. The corpus
adapter already exposes image-anchor inspection (`inspectImageAnchors` / `readFixtureImageAnchors`)
and package-part inspection (drawing/media/rel-id uniqueness), plus a streaming-sheet write
capability — jointly enough to assert a streaming-written image once the feature exists.

## Open questions

- API shape: require all images for a sheet before its commit, or addable only up-front at sheet
  creation?
- Absolute-anchor images (independent of row flush timing) vs cell-range anchors that reference
  not-yet-written rows.
- De-duplicate image media shared across sheets into a single media part?
- Memory posture: large image binaries must not defeat the streaming memory guarantee — they should
  stream to the package rather than be retained in full alongside all sheet buffers.

Related: `image-anchor-emu-from-real-column-geometry`, `header-footer-image-authoring`,
`streaming-write-memory-and-shared-strings-tradeoff`, `streaming-writer-worksheet-splice-rows-columns`.
