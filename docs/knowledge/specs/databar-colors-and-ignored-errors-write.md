# Data-bar colors and "ignored errors" suppression on write

Cluster: conditional-formatting / styles

## Scenario

Two write-side OOXML capabilities the library omits today:

1. **Data-bar colors.** When applying a data-bar conditional-formatting rule, a caller wants control
   over the bar `color` and, for bars that can extend negative, the `negativeFillColor`. Both are
   first-class attributes of the OOXML `<dataBar>` element. The current data-bar support writes the
   bar geometry (min and max cfvo) but not these colors, so every bar renders in the default color.
2. **"Number stored as text" suppression.** A caller who intentionally stores numeric-looking values
   as text (leading-zero IDs, long account numbers) wants to suppress the spreadsheet app's
   green-triangle "number stored as text" warning for those cells, by emitting an `<ignoredErrors>`
   declaration into the worksheet. The library produces no such element today.

> Spec note, not a corpus case: this bundles two distinct write-side features that need API design, not
> a single bug with a clean reproduction. The known patch is a crude blanket hack (see caution); the
> durable value is the desired capability and the design constraints. Each becomes assertable, via
> inspectPackage worksheet-XML facts plus a round-trip onto the rule, once the authoring API exists.

## Desired behavior

- **Data-bar `color` and `negativeFillColor` are first-class.** A data-bar rule accepts a bar `color`
  and a `negativeFillColor`, expressed with the same color model used elsewhere (argb, or theme plus
  tint). On write they serialize as `<color>` and `<negativeFillColor>` children of `<dataBar>`; on
  read they round-trip back onto the rule. The public type surface exposes them precisely. Model the
  **whole** data-bar color set now, including `negativeBarColorSameAsPositive`, `axisColor` and border
  colors, to avoid a second breaking pass. This ties to the databar round-trip case, whose `gradient`
  flag is a known-open.
- **The extended data-bar color children emit in schema order, negative FILL before negative
  BORDER.** The richer color set lives in the `x14` data-bar extension (`<x14:dataBar>` under the
  worksheet `extLst`), whose child sequence is schema-fixed. In particular `<x14:negativeFillColor>`
  must be emitted **before** `<x14:negativeBorderColor>`; a writer that emits them in the reverse
  order, or in caller-supplied order, produces a package that spreadsheet applications reject as
  corrupt and refuse to open. The writer must order the extension's color children by the schema,
  independent of the order in which the caller set the properties, the same fixed-child-order
  discipline called out for `<ignoredErrors>` below and for the streaming writer's trailing elements.
  Today's legacy writer emits no `x14` data-bar extension at all, so the corruption is latent until
  the negative-color authoring API exists; it must land correctly-ordered by construction.
- **Targeted, opt-in ignored-errors.** A worksheet can carry an `<ignoredErrors>` block whose
  `<ignoredError>` entries name a range (sqref) and the error categories to ignore
  (`numberStoredAsText`, `evalError`, `formula*`, `emptyCellReference`, `listDataValidation`, …). The
  caller declares, per worksheet or per range, which category is ignored; the writer emits the
  corresponding element. **Nothing is emitted by default**, so today's behavior of writing no
  ignoredErrors is preserved unless the caller opts in.
- **No blanket suppression.** The naive approach, one `sqref="A1:XFD1048576" numberStoredAsText="1"`
  covering the whole grid, is rejected: it silences the warning everywhere, including genuine
  mistakes, and hard-codes the grid dimensions. The API names the specific ranges.

## Open questions

- API shape: a worksheet-level `ignoredErrors` collection of `{ranges, categories}` (the sqref packs
  multiple ranges, so a category-to-range-set map is the natural model), or a per-range convenience
  helper? Both, with the helper sugaring the collection?
- Element ordering: `<ignoredErrors>` has a fixed position in the worksheet element sequence, so the
  writer must place it correctly relative to neighbors. This is the same trailing-element-order
  discipline the streaming writer currently gets wrong for hyperlinks, CF and dataValidations.
- Data bars: confirm the full negative-bar, axis and border color set is modeled in one pass rather
  than just `color` plus `negativeFillColor`.

Related: `databar-conditional-formatting-roundtrip`, `conditional-formatting-cellis-and-expression-semantics`,
`public-type-surface-matches-runtime`, `numeric-string-preserved-not-coerced`.
