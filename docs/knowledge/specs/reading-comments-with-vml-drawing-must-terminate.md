# Reading cell comments backed by a legacy VML drawing must terminate

Cluster: comments / security

## Scenario

A workbook produced by a third-party, non-Excel generator carries cell comments and notes the
classic way: a `comments{N}.xml` part listing each note's author and text runs, plus a companion
legacy `vmlDrawing{N}.vml` container, the shape-per-note file the notes anchor into. A user opens
the workbook simply to read its data. Instead of returning, the read **never terminates and the
process hangs**. It does not crash or return wrong data; it stalls. The reproduction is a real
Chinese-locale template with ~23 hidden note shapes, all `ObjectType="Note"`, a comments part with
matching author and text runs, and standard `MoveWithCells` and `SizeWithCells` anchors.

> Captured as a spec note, not a corpus case: a read that **hangs** must never enter the behavior
> corpus, since it would stall CI. The requirement, bounded and terminating parsing of the comment
> plus VML pair, is recorded here; a bounded-time regression check belongs in a dedicated perf or
> security harness with a hard timeout. Confirmed: the current reader still times out, over 30s, on
> the fixture.

## Desired behavior

- Reading a workbook whose comments are backed by a legacy VML drawing **completes promptly** and
  yields the worksheet. The mere presence of a `vmlDrawing` part alongside a `comments` part must
  never stall parsing.
- The read is reported as success, with no error, for this schema-valid foreign-generated file.
- Each cell's note text, from the comments part, is exposed on the corresponding cell after the
  read.
- A round-trip preserves the comments part and its companion VML drawing part with unique
  relationship ids, the same unmodeled-part-passthrough principle as
  `header-footer-image-survives-roundtrip` and the pivot and chart notes.

## Prior art / root cause

The stall is a parsing pathology in the legacy-VML and comments path, not intrinsic to the format,
since Excel and other readers open the same file instantly. The likely culprits, to verify during the
rewrite: a VML tokenizer that does not make monotonic progress on this generator's shape and markup
style, meaning an unbounded or re-scanning loop, and cross-referencing between comments and VML that
degrades on the hidden-note anchors. Because the VML parser consumes attacker-influenced markup, the
fix must be framed as a hostile-input concern: bounded work per part, guaranteed forward progress,
and a hard ceiling consistent with the fork's zip-bomb and unbounded-allocation stance.

## Open questions

- Is the hang in the VML tokenizer specifically, or in assembling the comment model from the VML
  anchors? A minimal reduction of the fixture would localize it.
- Should an unparseable or pathological VML drawing degrade gracefully, surfacing the comments without
  the shape geometry, rather than fail the whole read?
- Where does the terminating-parse guarantee get asserted? A perf or security harness with a
  wall-clock cap on this fixture, since a durable corpus case cannot host a potentially-hanging read.

Related: `whole-column-data-validation-bounded-memory`, `bounded-memory-large-workbook-read`,
`tolerant-parse-unclosed-vml-tags` (VML robustness), `empty-comments-read-as-blank-notes`.
