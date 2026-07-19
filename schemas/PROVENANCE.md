# Provenance — vendored OOXML schemas

## Source

- **Standard:** ECMA-376, *Office Open XML File Formats*
- **Part:** Part 4 — *Transitional Migration Features*, 5th edition (December 2016)
- **Publisher:** Ecma International — <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
- **Retrieved:** 2026-07-19
- **Distribution archive:** `ECMA-376-4_5th_edition_december_2016.zip`
  - SHA-256: `bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd`
  - Inner archive `OfficeOpenXML-XMLSchema-Transitional.zip`
    SHA-256: `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`

The 26 `.xsd` files in `ooxml-transitional/` are extracted **verbatim and
unmodified** from that inner archive. Do not hand-edit them — if a correction is
ever needed, re-extract from the source archive so provenance stays intact.

## Licensing

Ecma International makes ECMA-376 and its schemas **freely available** and
permits their reproduction. The Office Open XML schemas are additionally covered
by Microsoft's **Open Specification Promise (OSP)**. They are redistributed here
unmodified, as development-time reference material, under those terms. This does
not alter the license of this project's own source code (see `../LICENSE`).

## Not vendored (and how to get it)

- **Strict schema set** — ECMA-376 Part 1, 5th edition (Dec 2016),
  `ECMA-376-1_5th_edition_december_2016.zip`
  (SHA-256 `9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129`),
  inner `OfficeOpenXML-XMLSchema-Strict.zip`. Add under `schemas/ooxml-strict/`
  only if a Strict-conformance path is implemented.
- **Open Packaging Conventions** (`[Content_Types].xml`, `.rels`) — ECMA-376
  Part 2. Small and stable; extract from Part 2's schema archive if ever needed.

## Reproduce

```sh
curl -sLO https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
unzip -p ECMA-376-4_5th_edition_december_2016.zip OfficeOpenXML-XMLSchema-Transitional.zip > t.zip
unzip t.zip -d schemas/ooxml-transitional
```
