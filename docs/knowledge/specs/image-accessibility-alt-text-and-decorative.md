# Image accessibility: title, alt-text description, and a decorative flag

Cluster: images

## Scenario

When a user embeds an image, they may need to attach accessibility metadata so screen readers and
compliance tooling, including Excel's own accessibility checker, can describe it: a **title**, a
longer **description (alt-text)**, and a flag marking the image as purely **decorative** so assistive
technology skips it. Today the image API models only the binary and its anchor, with no way to
set any of this, so the metadata cannot be authored and is absent from the output.

> Spec note, not a corpus case: probing shows the image API (`addImage` and anchor) has no title,
> description or decorative parameter, so the metadata cannot be authored to assert a round-trip. The
> feature does not exist. The durable value is the desired API and its OOXML mapping; promote to
> a corpus case once an authoring API lands.

## Desired behavior

- An image can carry a **title** and a **description (alt-text)**, and both are preserved across a
  write/read round-trip.
- These appear as attributes on the picture's **non-visual drawing properties** in the emitted
  drawing XML: `<xdr:cNvPr>` carries `title` and `descr` attributes, where today it emits only `id`
  and `name`.
- An image can be marked **decorative**; the decorative flag is emitted into the drawing's extension
  list, the `<a:extLst>` decorative marker on the shape's non-visual properties, and survives a
  round-trip.
- Reading a foreign file with these attributes surfaces them on the image model, and a round-trip of
  an image that already carries them does not drop them.

## Open questions

- API shape: options on `addImage` (`{ title, description, decorative }`) against setters on the
  returned image handle. Reuse the existing image and anchor vocabulary.
- Interaction: a decorative image typically should not also carry a description, since Excel's checker
  treats "decorative" and "has alt-text" as mutually exclusive. Validate, warn, or leave to the
  caller?
- Should the same title and description generalize to other drawing objects (shapes, charts) via
  the shared non-visual-properties model rather than an image-local one?
- Scope of the decorative extension: emit only when set, defaulting off, preserving today's byte
  output for images without accessibility metadata.

Related: `preserve-drawing-shapes-on-roundtrip`, `image-range-anchor-edit-as-mode-honored`,
`comment-note-box-fits-its-text`.
