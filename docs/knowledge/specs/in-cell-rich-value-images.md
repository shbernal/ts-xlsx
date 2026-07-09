# Reading in-cell ("Place in Cell") rich-value images

Cluster: images

## Scenario

Excel's "Place in Cell" feature stores an image *inside a cell value* rather than as a
floating drawing anchored over the grid. This uses the rich-value mechanism: the cell carries a
value-metadata (`vm`) index that points through the workbook metadata part to a rich-value record,
which references the image via a rich-value relationship to a media part. The legacy `#VALUE!`
error is stored in the cell as a fallback for applications that don't understand rich values.

A reader that only understands drawing-anchored images (twoCell/oneCell anchors) surfaces none of
these: `getImages()` returns nothing, and the image-bearing cells look empty or like errors. The
image is silently invisible to the caller.

## Desired behavior

- Reading a workbook whose cells use in-cell rich-value images surfaces those images, each
  **associated with its host cell** (e.g. C2), not with a drawing-anchor position.
- Resolution follows the real chain: cell `vm` index → workbook metadata → rich-value record →
  rich-value relationship → media part, yielding the correct image bytes for each cell.
- The legacy `#VALUE!` error fallback stored in the cell must not mask the presence of the in-cell
  image — the model should expose the image, and treat the error as the documented fallback.
- This is distinct from drawing-anchored images: the API should make clear which images are
  in-cell (part of a cell's value) vs. floating (anchored over the grid).

## Open questions

- The public model shape for an in-cell image: a distinct cell value kind (e.g. an image value)
  vs. a parallel per-cell image accessor. It must round-trip.
- Write-side authoring of in-cell images (create a "Place in Cell" image) — in scope, or read/
  preserve only for now?
- Hostile-input hardening: the metadata → rich-value → media indirection is attacker-influenced;
  bound the indirection depth and validate indices so a crafted file cannot force unbounded work
  or dangling-reference crashes.
- Interaction with `getImages()` (drawing-anchored) so callers can enumerate both kinds without
  conflating them.

Related: `image-by-filename-is-node-only`, `image-anchor-emu-from-real-column-geometry` (the
drawing-anchored image model this extends beyond).
