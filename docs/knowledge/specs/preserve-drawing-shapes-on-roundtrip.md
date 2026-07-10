# Preserve drawing shapes (autoshapes, text boxes) on round-trip

Cluster: images

## Scenario

A user starts from a template workbook containing drawing shapes — autoshapes, text boxes,
connectors, other vector graphics anchored to the sheet (distinct from raster images). They edit cell
values and write it back. On reopening, all cell data survives but every shape has vanished: the
drawing layer the library does not model is silently dropped during read/modify/write. Content the
user never touched should survive a round-trip unchanged.

> Spec note, not a corpus case: the library has no shape model at all, so this is a feature-gap /
> data-loss-prevention requirement rather than a discrete assertable bug. The durable value is the
> preservation contract (and the reconciliation hazard where a drawing mixes modeled images with
> unmodeled shapes). It belongs with the package-part-preservation family (comments, VML, tables).

## Desired behavior

Two levels of ambition:

1. **Preservation (minimum, higher value).** Drawing parts, their shapes (`xdr:sp` autoshapes, text
   boxes, connectors), and associated rels/media are treated as **opaque pass-through** when the
   library does not fully model them, so untouched content round-trips structurally intact. This
   aligns with the fork's "don't silently destroy user data" stance. **Care at the mixed-content
   edge**: when a drawing part mixes modeled images with unmodeled shapes, the writer must **merge**
   regenerated image anchors with preserved shape anchors, not choose one wholesale — otherwise
   regenerating the drawing from the image model alone drops the shapes (the current failure), or
   preserving verbatim drops newly-added images.
2. **First-class modeling (stretch).** A real API to create/edit shapes as siblings of images in the
   drawing collection. The concrete authoring surface (from a second, more detailed report):
   `addShape(shapeSpec, rangeOrAnchor, { hyperlink, tooltip })` where the spec carries a **preset
   geometry** (`prstGeom` presets — `roundRect`, `rect`, `ellipse`, arrows, …), an **anchor** (one- or
   two-cell), a **rotation**, a **solid fill** (RGB, later scRGB/theme), an **outline** (`ln`: weight,
   color, dash pattern like `sysDash`), and a **text body** — an ordered list of paragraphs, each with
   a horizontal alignment and a list of runs (plain text, or text + font: bold/italic/size/color),
   plus the box's vertical alignment. OOXML mapping: shapes are `<xdr:sp>` in a one/two-cell anchor,
   parallel to `<xdr:pic>`; `spPr` holds `xfrm`/`prstGeom`/`solidFill`/`ln`, `txBody` holds `bodyPr`
   (vertical align) + `a:p` paragraphs (`a:pPr` align, `a:r` runs with `a:rPr`+`a:t`), `nvSpPr` carries
   name/id + optional hyperlink. Reuse the existing image-anchor and color/theme vocabulary rather
   than a shape-local one. Large surface; verbatim preservation (level 1) is the safer near-term win,
   with typed modeling layered on where fidelity is guaranteed. Unknown geometry presets, unsupported
   fill/line types, and unmodeled run properties must degrade predictably (prefer verbatim
   passthrough) rather than fail the load.

## Prior art

A long-standing, repeatedly-reported gap hit independently by template-driven workflows. It
co-occurs with the general "foreign-generator content dropped on round-trip" family the fork already
tracks via package-part preservation (comments, VML, tables, pivots).

## Open questions

- Opt-in or default? The fork principle argues **default-preserve** for content the user did not touch.
- Mixed drawing part reconciliation: anchor-level merge vs whole-part passthrough when any unmodeled
  element is present.
- Fidelity contract for preserved parts: byte-identical, or structurally-equivalent-after-normalization.
- Do shape anchoring semantics (twoCell/oneCell/absolute) need to be understood even for pure
  passthrough, or can offsets be preserved verbatim?

Related: `roundtrip-preserves-unmodeled-package-parts`, `load-workbook-with-chart-drawing-does-not-crash`,
`embedded-chart-read-write`, `image-range-anchor-edit-as-mode-honored`.
