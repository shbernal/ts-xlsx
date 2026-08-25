# Excel decodes `_xHHHH_` in cell text, and does it in one left-to-right pass

Cluster: xlsx-io

## Scenario

XML 1.0 has no syntax for a C0 control character. `&#1;` is not an escape for U+0001; it is another
spelling of the same illegal document. SpreadsheetML's answer is an in-band convention, `_xHHHH_`,
where `HHHH` is four hex digits naming a code point. A writer that wants to carry such a character
in a cell has to use it, and a reader that wants to read Excel's files has to undo it.

That much is widely implemented (POI and openpyxl both do it) and widely repeated. It is still a
claim about what an *application* does, not what a schema says, and `ST_Xstring` types a sheet name
exactly as it types a `<t>`, so the schema cannot even tell you where the convention applies. The
questions that actually decide an implementation are narrower than "does Excel decode it":

- Does it decode in the shared-strings pool, in an inline `<is>`, or both?
- Does it decode in the cached `<v>` of a `t="str"` formula cell, which is a value but not a `<t>`?
- What happens to `_x005F_x0041_`, the escaped underscore in front of an escape? One pass yields the
  literal text `_x0041_`; two passes collapse it to `A`. Both are defensible from the prose.
- Is the decode restricted to characters XML could not have carried anyway?
- How strict is the grammar? Does `_x041_` or `_xZZZZ_` decode to anything?

## Measured behavior

Excel Desktop (Microsoft 365, version 16.0, build 20228, Windows), driven headless over COM with
`Calculation` forced to manual so a recalculation could not overwrite a cached `<v>` before it was
read. A package written by this library, then textually patched so the escape sequences reach the
parts verbatim rather than being escaped by our own writer; validated clean by `ooxml-validate`
before opening. Values read through `Range.Value2` and reported as character codes, so nothing is
lost to console rendering.

| Cell | On disk | Read back |
| --- | --- | --- |
| A1 | `<c t="inlineStr"><is><t>_x0041_</t></is></c>` | `A` |
| A7 | `t="s"` into `<si><t>_x0041_</t></si>` | `A` |
| B1 | `<c t="str"><f>A1</f><v>_x0041_</v></c>` | `A` |
| A2 | `<t>_x005F_x0041_</t>` | `_x0041_` (7 characters) |
| A3 | `<t>_xZZZZ_</t>` | `_xZZZZ_` |
| A4 | `<t>_x041_</t>` | `_x041_` |
| A5 | `<t>_x00041_</t>` | `_x00041_` |
| A6 | `<t>a_x0009_b</t>` | `a`, U+0009, `b` |

Five facts fall out.

**The pool and the inline string behave identically.** A1 and A7 are the same escape reaching the
cell by different routes, and both decode. So the convention belongs to the `<t>`, not to the part
it happens to live in.

**The cached result of a string formula decodes too.** B1 is the case the schema cannot settle and
the one most likely to be assumed wrong in either direction. It decodes, so a `t="str"` `<v>` is
cell text for this purpose despite not being a `<t>`.

**The decode is one left-to-right pass.** A2 is the proof: `_x005F_` yields `_`, scanning resumes
past it on `x0041_`, and with no leading underscore left that is ordinary text. Excel reads the cell
as the literal `_x0041_`. A decoder that resolves `_x005F_` in a pass of its own, before or after
the rest, reads that cell as `A` and silently destroys every value that legitimately looks like an
escape.

**The grammar is closed and exact.** Four hex digits, no more and no fewer, between `_x` and `_`.
A3, A4 and A5 are returned untouched, so a reader must not be liberal here: near-misses are ordinary
text and stay that way.

**The decode is unconditional, not a repair.** A6 carries a tab, which XML 1.0 permits literally, and
Excel decodes it anyway. A reader that only undid the escape for characters XML cannot represent
would disagree with Excel about files Excel itself wrote.

## What follows for this library

Both directions are in `src/xml/`: `escapeSpreadsheetText` in `xml.ts`, `decodeSpreadsheetText` in
`xml-read.ts`. They are inverses, and the corpus case
`escaped-characters-in-cell-text-decode-on-read` locks that with the package above as its fixture.

Where the convention applies is a separate decision the writer makes, and the measurement above does
not settle it: it says the escape *works* in cell text, not that it works anywhere else. This library
escapes cell text and refuses everything structural. The module header of `src/xml/xml.ts` states
the line and why it falls there.

## Provenance

`source: excel-desktop-verification`, the tier ADR-0013 describes. One Excel build, one host.

The COM tier is sufficient here: every fact above is a cell value, which `Range.Value2` reports
directly, so nothing needed the rendering tier. Reproducible from the repo: the probe package is
committed as the corpus fixture
`test/corpus/fixtures/escaped-characters-in-cell-text-decode-on-read/xhhhh-escapes.xlsx`, and
re-running it is a matter of opening that file and reading A1:A7 and B1.

Related: `excel-repair-on-open-structural-constraints`.
