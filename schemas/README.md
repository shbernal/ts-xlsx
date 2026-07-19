# OOXML schemas — vendored reference

Machine-readable ground truth for the file format this library reads and writes.
These are the **XML Schema (XSD) definitions** from ECMA-376 (Office Open XML),
vendored verbatim so that any agent — human or AI — implementing a parser or
writer can consult the authoritative element structure, attribute types,
enumerations, and child-element ordering **without a network round-trip**.

## What is here

- [`ooxml-transitional/`](./ooxml-transitional/) — the complete **Transitional**
  conformance schema set (26 `.xsd` files) from ECMA-376 Part 4, 5th edition.
  Transitional is what Microsoft Excel actually emits, so it is the set that
  matters for reading real-world `.xlsx` files. See [`PROVENANCE.md`](./PROVENANCE.md)
  for source, edition, hashes, and licensing.

The Strict set (ECMA-376 Part 1) is intentionally **not** vendored — real files
are overwhelmingly Transitional. `PROVENANCE.md` records how to obtain it if a
Strict-conformance path is ever needed.

## The files that matter for `.xlsx`

`.xlsx` is a spreadsheet, so start from `sml.xsd` and follow its imports. The
relevant closure:

| File | Namespace / role |
| --- | --- |
| `sml.xsd` | **SpreadsheetML** — worksheets, cells, styles, shared strings, tables, pivot, defined names. The core. |
| `shared-commonSimpleTypes.xsd` | Shared simple types (`ST_*`) referenced everywhere. |
| `shared-relationshipReference.xsd` | The `r:id` relationship-reference attributes. |
| `dml-spreadsheetDrawing.xsd` | DrawingML anchoring of images/charts onto a sheet (`xdr:` namespace). |
| `dml-main.xsd` | DrawingML core (shapes, fills, colors, transforms). |
| `dml-chart.xsd`, `dml-chartDrawing.xsd` | Charts embedded in a workbook. |
| `dml-picture.xsd`, `dml-lockedCanvas.xsd`, `dml-diagram.xsd` | Pictures, locked canvas, SmartArt. |
| `vml-main.xsd`, `vml-spreadsheetDrawing.xsd`, `vml-officeDrawing.xsd` | Legacy **VML** — Excel still uses it for comment/note shapes and some form controls. |
| `shared-documentProperties*.xsd` | `docProps/core.xml`, `app.xml`, `custom.xml`. |

The WordprocessingML (`wml.xsd`) and PresentationML (`pml.xsd`) schemas are part
of the set because the VML and DrawingML schemas import shared types from them;
they are otherwise irrelevant to a spreadsheet library.

Not covered by this set (they live in ECMA-376 **Part 2**, the Open Packaging
Conventions): `[Content_Types].xml` and the `.rels` relationship-part schema.
Those are small and stable; see `PROVENANCE.md` for where to find them.

## These are for reference, not validation

Conformance validation is already handled by an independent oracle — the
repository-owned .NET tool wrapping Microsoft's `OpenXmlValidator`
(`pnpm run validate:ooxml`, see [ADR-0002](../docs/decisions/0002-ooxml-validation-oracle.md)).
**Do not** wire these XSDs into a second validation path. Their job is to be read.

## Not shipped

This directory is repo-only reference material. It is excluded from the published
npm package by the `files` allowlist in `package.json` (`dist`, `LICENSE`,
`README.md` only). It adds **zero** weight to what users install.
