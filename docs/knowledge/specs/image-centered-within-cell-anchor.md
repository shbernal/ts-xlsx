# Center an image within a cell or range, preserving aspect ratio

Cluster: images

## Scenario

A user wants to place an image inside a cell (or a small range) so it appears **centered** within that
region with margins around it — not pinned to the top-left corner and not stretched to fill. The two
placement modes available today are: a top-left anchor plus an explicit pixel extent (image at the
cell's top-left, overflowing into neighbors, no centering), and a two-cell range anchor (image
stretched to exactly fill the range, distorting aspect ratio). Neither centers with padding.

> Spec note, not a corpus case: a convenience placement mode that does not exist yet. It is
> expressible over the existing OOXML anchor primitives, so it is a higher-level helper, not a schema
> change; the durable value is the placement contract and the enabling geometry.

## Desired behavior

- An image-placement mode that **centers** an image within a target cell/range, preserving the
  image's aspect ratio and leaving symmetric margin around it — "put this image in B2:D6, centered,
  without distorting it."
- The library computes the image's intrinsic size (or a caller-supplied size/scale), measures the
  target region's pixel dimensions from the involved **column widths and row heights**, and derives a
  `oneCellAnchor` (from-cell + inward `colOff`/`rowOff` offsets + `ext`) so the rendered image is
  centered with equal horizontal/vertical padding.
- **Aspect-ratio-preserving fit-within (letterbox)** is the natural default; stretch-to-fill remains
  available (today's two-cell behavior).
- Built on existing anchor vocabulary and the column-width/row-height → pixel → EMU conversion already
  needed for pixel-accurate layout — no new schema.

## Open questions

- API shape: a placement enum on the anchor (`top-left | fill | center-fit`), or a distinct helper
  that computes offsets from a range + desired size? Reuse the existing image-anchor vocabulary.
- Sizing source: intrinsic dimensions vs caller width/height vs scale factor; behavior when the image
  is larger than the region (shrink-to-fit vs overflow).
- Interaction with move/size-with-cells (`editAs`): centering is computed **at write time** from the
  then-current column/row sizes and is not dynamic — document that it does not re-center if the user
  later resizes rows/columns.
- Should centering also work for ranges spanning merged cells?

Related: `image-anchor-emu-from-real-column-geometry`, `image-anchor-fractional-offset-respects-cell-size`,
`image-range-anchor-edit-as-mode-honored`, `fractional-image-anchor-positioning`.
