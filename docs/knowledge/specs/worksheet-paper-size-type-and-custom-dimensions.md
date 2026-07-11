# Worksheet paper size: complete type coverage and custom dimensions

Cluster: page-setup

## Scenario

Print/page setup exposes a paper size. Two distinct gaps show up around it.

First, the **type surface is incomplete**. The public `paperSize` type enumerates only a subset of the
OOXML paper codes, omitting common ones — A3 (code 8) is the headline example, but the standard defines
dozens (A3, A5, B4, B5, envelope and rotated variants, and so on). The underlying numeric code almost
certainly serializes fine when written, so this is not a runtime bug; it is a contract defect. Because
this fork is TypeScript-first and treats the types as the primary documentation, a `paperSize` the
docs list but the type rejects is a real failure: a caller who writes `paperSize: 8` for A3 gets a
false compile error, or reaches for `as any`.

Second, some real print layouts need a **custom paper size** — an explicit width/height rather than one
of the enumerated standard codes (large-format plots, receipt/label rolls, non-ISO regional stock).
Today only the enumerated codes are expressible, so those layouts cannot be authored at all.

> Spec note, not a corpus case: the type-coverage half is a type-surface requirement (enforced by a
> type-level test, not a runtime assertion), and the custom-dimensions half is a new capability with no
> current behavior to baseline. Existing page-setup round-trip behavior (fit-to-page, orientation,
> margins) is already locked by corpus cases; this note is the umbrella policy for the paper-size axis.

## Desired behavior

- **Complete, honest `paperSize` typing.** The published type must admit every OOXML paper size code
  the writer can emit — not a curated subset. Prefer deriving the accepted set from the format's
  defined codes (a named enum or a documented numeric union) so a code the format allows can never be a
  compile error. A type-level test asserts representative codes (A3 = 8, and a spread of the higher
  codes) are assignable, so the surface cannot silently regress to a subset again. If a friendly named
  form is offered (e.g. `'A3'`), it maps to the numeric code; the raw numeric code stays accepted for
  forward compatibility with codes we have not named.

- **Custom paper dimensions.** A caller can specify an explicit page width and height (with a unit —
  the OOXML `pageSetup` `paperWidth`/`paperHeight` attributes take a length with a unit suffix, e.g.
  millimetres or inches) instead of, or overriding, an enumerated `paperSize`. When custom dimensions
  are set, they must serialize to `paperWidth`/`paperHeight` and round-trip faithfully; reading a file
  that already carries custom dimensions must preserve them rather than snapping to the nearest
  enumerated size. Define precedence when both a `paperSize` code and explicit dimensions are present
  (OOXML: the explicit width/height override the code).

## Open questions

- Naming: keep a single `paperSize` field accepting `number | 'A3' | … | {width, height, unit}`, or
  split the enumerated code and the custom-dimension object into distinct fields to avoid a union that
  is awkward to narrow?
- Unit handling for custom dimensions: accept a unit suffix string as OOXML stores it, or take a number
  plus a separate unit enum and format the suffix on write?
- Do we ship the full named-code enum (dozens of entries, most rarely used) or only name the common
  ones and leave the rest as documented numeric codes?
- Reading an unknown/rare numeric code: surface it verbatim (forward-compatible) rather than dropping
  or normalizing it.

Related: `pagesetup-fit-to-page-round-trips`, `column-width-and-pagesetup-roundtrip-fidelity`,
`page-margins-must-be-complete`, `public-type-surface-matches-runtime`.
