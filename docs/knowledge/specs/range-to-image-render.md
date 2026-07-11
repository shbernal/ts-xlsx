# Render a cell range as a styled image ("copy as picture")

Cluster: images

## Scenario

A user wants to take a rectangular range of cells (e.g. `A1:C15`) and produce a raster or vector
image that visually reproduces those cells as a spreadsheet application would render them — cell
values, number formatting, fonts, colors, fills, borders, alignment, merged cells, and column
widths / row heights. This mirrors the desktop "Copy as picture" affordance and is used to embed
spreadsheet snippets into reports, emails, dashboards, or web pages without shipping the underlying
data.

> Spec note, not a corpus case: this is a substantial new rendering capability with no current
> behavior to baseline, and a heavy design decision (a layout/rasterization surface, its dependency
> footprint, and its security posture). Recording the desired contract and the camp trade-offs feeds
> a Phase 3 decision; it is not something to assert against today's code.

## Desired behavior

- **Range → image API.** Take a worksheet plus a range reference (or range object) and return an image
  (PNG/SVG buffer or a stream) rendering those cells with their applied styling — shape roughly
  `render(worksheet, "A1:C15", { format: "png" | "svg", scale })`.

- **Fidelity floor.** The rendering must honor at minimum: cell display values (post-number-format,
  including dates/currencies/percentages — consumes the displayed-value work of
  `cell-value-raw-and-displayed-accessor`), font (family/size/weight/italic/color), fill/background
  color, borders (style + color), horizontal/vertical alignment and text wrap, merged-cell regions
  spanning the correct area, and geometry (column widths and row heights translated to pixels via the
  same real-geometry mapping used for image anchoring, `image-anchor-emu-from-real-column-geometry`).

- **Security- and dependency-conscious.** Two ecosystem camps exist: (1) translate the range to an
  HTML table + CSS and rasterize via a headless browser — reuses a mature layout engine but adds a
  heavy runtime dependency and a large trusted-input attack surface; (2) draw directly to a
  canvas/SVG surface from the parsed style model — keeps the dependency tree small (aligned with this
  fork's supply-chain stance) but reimplements text layout, wrapping, and border/merge geometry. The
  fork's minimal-dependency and hostile-input stances favor camp (2), at least for a deterministic SVG
  output that needs no native binaries; a PNG rasterizer can be an optional add-on.

## Open questions

- Output formats: SVG (vector, deterministic, no native deps) as the baseline, PNG (needs a
  rasterizer) as opt-in? Start with SVG to stay dependency-light.
- How much text-layout fidelity to promise — exact wrapping/overflow/shrink-to-fit is a deep rabbit
  hole; define a documented subset and its limits rather than implying pixel-perfect parity with a
  specific application.
- Shared surface with `worksheet-to-html-export`: HTML export and image render both need the same
  style→visual translation; build the style-to-visual model once and target both HTML and SVG from it.
- Whether font metrics require embedding/measuring real fonts (for accurate width) or an approximation
  suffices for the common case.

Related: `worksheet-to-html-export`, `cell-value-raw-and-displayed-accessor`,
`image-anchor-emu-from-real-column-geometry`, `indexed-palette-colors-resolve-to-concrete`,
`minimal-audit-clean-dependency-tree`, `no-unsafe-eval-csp-compatible`.
