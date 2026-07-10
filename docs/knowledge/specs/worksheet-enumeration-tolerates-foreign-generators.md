# Worksheet enumeration and lookup must work on foreign-generated files

Cluster: security/deps

## Scenario

A `.xlsx` produced by a non-Excel generator (a cloud spreadsheet service, a server-side exporter) is
opened for reading. The caller looks up a worksheet by its declared name and gets `undefined` back;
the workbook appears to contain zero worksheets, so the next method call throws a downstream
`Cannot read property … of undefined`. Opening the same file in desktop Excel and pressing Save —
which rewrites the package, normalizing its structure and adding a few KB — makes it read correctly
afterward. Two independent reporters describe the same failure. The reader is being too strict about
some structural convention the foreign generator does not follow (relationship id casing, part-name
path shape, ordering, or an omitted-but-defaultable attribute), and silently ends up with an empty
worksheet set instead of the sheets that are plainly present.

> Spec note, not a corpus case yet: pinning it needs a real foreign-generated fixture that reproduces
> the empty-enumeration failure (none was attached), and the true trigger must be identified rather
> than guessed — a synthetic file I hand-craft may not match the real structural quirk. The durable
> value is the tolerance requirement and the failure signature; promote a fixture and assert
> non-empty enumeration once a genuine reproducer is in hand.

## Desired behavior

- **Worksheet enumeration reflects the sheets the package actually declares.** Reading a foreign but
  spec-valid workbook exposes a non-empty worksheet list whose names match the workbook part, without
  a round-trip through Excel first.
- **Lookup by declared name returns a real worksheet**, not `undefined`, for every sheet the workbook
  part lists. A name that genuinely is not present returns a clear absence the caller can test — never
  a value that throws two calls later.
- **The reader tolerates benign structural variation** other producers legitimately emit: relationship
  id / target casing and path shape, element ordering, and attributes that have a spec default when
  omitted. It binds sheets to their parts by the relationship graph, not by assuming Excel's exact
  serialization.
- **A structurally broken package fails loudly, at load, with context** (which part/relationship could
  not be resolved) — never by degrading to a silently empty workbook that surfaces as an `undefined`
  crash in the caller's code. Robustness here must not become unbounded trust of hostile input; it is
  tolerance of valid-but-non-Excel structure, with malformed input still rejected cleanly.

## Open questions

- The actual trigger(s): sheet-to-rId binding assumptions, worksheet part-path patterns
  (`xl/worksheets/…` variants), or a missing defaultable attribute — needs a real foreign fixture to
  pin, ideally more than one generator.
- Whether a lenient "recover the sheets by scanning parts" fallback is worth having when the
  relationship graph is malformed, or whether that crosses into trusting broken input.
- How much of this the namespace/BOM-tolerant reading work already covers versus what is specific to
  the workbook→worksheet relationship resolution.

Related: `namespace-agnostic-bom-tolerant-ooxml-reading`, `defined-names-tolerate-non-address-tokens`,
`load-workbook-with-chart-drawing-does-not-crash`, `bounded-memory-on-load`.
