# Distinguish a cell-embedded ("in-cell") image from a floating anchored drawing

Cluster: images

## Scenario

A user adds an image and anchors it to a cell range (say C2:E6). On export the image floats *over* the
cells as an independent drawing object: it is not owned by any single cell, is not returned by
cell-value reads, and is not sorted/filtered/moved with the row it visually sits on. The user expected
the picture to be embedded *in* the cell — to behave as the cell's value, live inside the cell's
bounds, and travel with the cell through sort, filter, and row insert/delete. These are two
fundamentally different features, and today only the floating one exists.

> Spec note, not a corpus case: true in-cell images are a new capability (a rich-value / cell-metadata
> feature), not a bug in current behavior — there is nothing to baseline yet. The floating-anchor
> behavior is already covered by the image-anchor cases. The durable value is naming the distinction
> and the design questions; a corpus case follows once (and if) rich-value in-cell images are
> implemented, at which point the adapter grows a capability to read an image-valued cell.

## Desired behavior

Distinguish the two ways an image can live in a sheet, and let callers choose:

- **Floating / overlay image (today's behavior).** A drawing object anchored to the sheet grid via a
  one-cell or two-cell anchor. It sits over the cells, is not a cell value, and — depending on its
  "move and size with cells" property (`editAs`) — may or may not reflow when rows/columns are
  inserted, deleted, or resized. `addImage(id, 'C2:E6')` produces this; the range only positions the
  overlay's anchors. This mode deserves explicit, documented coverage of the `editAs` /
  "move and size with cells" property, which is the usual source of "why does my picture float / not
  move with the cell" confusion.
- **In-cell / embedded image.** The image is the *value* of exactly one cell (Excel's "Place in Cell"
  / IMAGE-function rich-value cells). It is owned by and clipped to that cell, moves and resizes with
  it, participates in sort/filter/insert/delete like any value, and reads back as the cell's value. On
  the OOXML side this is not an `xdr:drawing` two-cell anchor at all — it is a rich value in the
  rich-data / cell-metadata parts (`xl/richData/*`, cell metadata `vm` attributes, and a rels-linked
  image), a distinct feature from classic drawings.

Prior art: classic implementations support only floating drawing anchors (one-cell "over" a cell or
two-cell "over" a range) with the `editAs` flag. Helpers that "add an image over a range" still produce
a floating two-cell overlay — not a true cell-embedded rich value. Genuine in-cell images require the
rich-value / cell-metadata machinery.

## Open questions

- Scope: commit to reading and writing true rich-value in-cell images (rich-data parts + cell
  metadata), or first just clarify/name the existing floating anchor modes and the "move and size with
  cells" behavior?
- API shape: how does a caller request an in-cell image vs a floating one — a distinct method, or an
  anchor option like `anchor: 'inCell'` targeting a single cell, rather than overloading the
  range-anchor call?
- Read-back: an in-cell image should surface as the owning cell's value; define the value shape (image
  ref + alt text) and how it coexists with a formula/number-format on the same cell.
- Fallback for older Excel lacking rich-value support (Excel writes a `#VALUE!`-style cached value plus
  the rich value): do we emit that compatibility shim?

Related: `image-range-anchor-edit-as-mode-honored`, `anchored-image-sppr-transform-detaches-in-libreoffice`,
`fractional-image-anchor-positioning`, `header-footer-image-authoring`, `public-type-surface-matches-runtime`.
