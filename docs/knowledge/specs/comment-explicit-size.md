# A cell comment can be given an explicit size

Cluster: styles

## Scenario

A user attaches a note (cell comment) and wants control over the popup box's dimensions — a fixed
width and height, so a short label gets a small box and a long annotation a large one, matching a
designed layout. Today the note's size is not author-controllable: the library emits a default box
and (once `comment-note-box-fits-its-text` lands) can grow the box to fit its text, but there is no
way to say "this note is 240×120 pixels." Users resort to post-processing the output zip to edit the
comment's VML drawing geometry by hand.

> Spec note, not a corpus case: this is an authoring-surface addition (accept explicit dimensions on
> a note) and a decision about units and precedence versus auto-fit — a design question, not a
> malformed-output bug. It becomes a corpus case once the API accepts a size and the emitted VML
> shape geometry is asserted.

## Desired behavior

- **A note accepts an explicit size** (width and height) alongside its text, and that size is written
  into the comment's VML drawing shape geometry so hosts render the box at the requested dimensions.
- **Explicit size and auto-fit are mutually exclusive, with a clear precedence**: setting an explicit
  size turns off `mso-fit-shape-to-text` (the auto-grow directive from
  `comment-note-box-fits-its-text`) for that note, because a caller who fixes the size does not want
  the host to override it. With no explicit size, auto-fit remains the default.
- **A well-defined unit**: the authoring API takes a documented unit (pixels are the most intuitive
  for callers; the VML geometry is emitted in the units the format requires, converted internally),
  so "240" means the same thing regardless of column widths or DPI.
- **Round-trips**: a note written with an explicit size reads back with that size, and re-saving
  preserves it.

## Open questions

- Authoring unit: pixels (convert to VML `pt`/EMU internally) versus exposing the raw format units.
  Pixels match user intuition and the screenshots in the requests.
- Anchor interaction: is the size absolute, or expressed as a cell-span (from/to anchor) the way
  floating shapes are? Absolute width/height is simpler and matches the request; a cell-span anchor
  is a possible richer alternative.
- Whether position (offset from the anchored cell) is part of the same API or a separate concern.

Related: `comment-note-box-fits-its-text` (already locks auto-grow-to-text),
`reading-comments-with-vml-drawing-must-terminate`.
