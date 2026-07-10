# Data-bar colors and "ignored errors" suppression on write

Cluster: conditional-formatting / styles

## Scenario

Two write-side OOXML capabilities the library omits today:

1. **Data-bar colors.** When applying a data-bar conditional-formatting rule, a caller wants control
   over the bar `color` and — for bars that can extend negative — the `negativeFillColor`. Both are
   first-class attributes of the OOXML `<dataBar>` element. The current data-bar support writes the
   bar geometry (min/max cfvo) but not these colors, so every bar renders in the default color.
2. **"Number stored as text" suppression.** A caller who intentionally stores numeric-looking values
   as text (leading-zero IDs, long account numbers) wants to suppress the spreadsheet app's
   green-triangle "number stored as text" warning for those cells, by emitting an `<ignoredErrors>`
   declaration into the worksheet. The library produces no such element today.

> Spec note, not a corpus case: this bundles two distinct write-side features that need API design, not
> a single bug with a clean reproduction. The known patch is a crude blanket hack (see caution); the
> durable value is the desired capability and the design constraints. Each becomes assertable (via
> inspectPackage worksheet-XML facts + a round-trip onto the rule) once the authoring surface exists.

## Desired behavior

- **Data-bar `color` and `negativeFillColor` are first-class.** A data-bar rule accepts a bar `color`
  and a `negativeFillColor`, expressed with the same color model used elsewhere (argb / theme+tint).
  On write they serialize as `<color>` and `<negativeFillColor>` children of `<dataBar>`; on read they
  round-trip back onto the rule. The public type surface exposes them precisely. Model the **whole**
  data-bar color set now — `negativeBarColorSameAsPositive`, `axisColor`, border colors — to avoid a
  second breaking pass. (Ties to the databar round-trip case, whose `gradient` flag is a known-open.)
- **Targeted, opt-in ignored-errors.** A worksheet can carry an `<ignoredErrors>` block whose
  `<ignoredError>` entries name a range (sqref) and the error categories to ignore
  (`numberStoredAsText`, `evalError`, `formula*`, `emptyCellReference`, `listDataValidation`, …). The
  caller declares, per worksheet or per range, which category is ignored; the writer emits the
  corresponding element. **Nothing is emitted by default** — today's behavior of writing no
  ignoredErrors is preserved unless the caller opts in.
- **No blanket suppression.** The naive approach — one `sqref="A1:XFD1048576" numberStoredAsText="1"`
  covering the whole grid — is rejected: it silences the warning everywhere (including genuine
  mistakes) and hard-codes the grid dimensions. The API names the specific range(s).

## Open questions

- API shape: a worksheet-level `ignoredErrors` collection of `{ranges, categories}` (the sqref packs
  multiple ranges, so a category→range-set map is the natural model), vs. a per-range convenience
  helper? Both, with the helper sugaring the collection?
- Element ordering: `<ignoredErrors>` has a fixed position in the worksheet element sequence; the
  writer must place it correctly relative to neighbors (this is the same trailing-element-order
  discipline the streaming writer currently gets wrong for hyperlinks/CF/dataValidations).
- Data bars: confirm the full negative-bar/axis/border color set is modeled in one pass rather than
  just `color` + `negativeFillColor`.

Related: `databar-conditional-formatting-roundtrip`, `conditional-formatting-cellis-and-expression-semantics`,
`public-type-surface-matches-runtime`, `numeric-string-preserved-not-coerced`.
