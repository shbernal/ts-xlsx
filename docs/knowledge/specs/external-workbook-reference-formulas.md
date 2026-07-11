# External-workbook reference formulas need package machinery, not a magic string

Cluster: formulas

## Scenario

A user wants a cell formula that pulls a value from a different workbook file — e.g. `Sheet1!A1` of a
sibling workbook `测试.xlsx` in the same folder. They set the cell formula to the literal text
`=[测试.xlsx]Sheet1!A1`. When opened in Excel or WPS, neither resolves the reference: Excel interprets
the bracketed name against its own default document path, and WPS only resolves its own
`=./[file.xlsx]Sheet1!A1` dialect. The reference does nothing because the library passed the string
through verbatim without generating any of the package machinery an external reference actually
requires.

> Spec note, not a corpus case: this is a capability the library does not model at all
> (cross-workbook references), not a reproducible bug with an implementation-blind assertion against
> today's code. It records the OOXML mechanism and the API design questions for Phase 3.

## Desired behavior

The library should author (and round-trip) formulas that reference cells/ranges in *other* workbook
files, such that Excel resolves them on open. In `.xlsx` an external reference is **not** the literal
`[filename.xlsx]Sheet1!A1` text inside `<f>` — it requires dedicated parts:

- An **externalLink part** (`xl/externalLinks/externalLink1.xml`) declaring the target workbook, its
  sheet names, and a cached snapshot of the referenced values (so the file shows a value when the
  target is offline).
- A **relationship** from that part to the target workbook, `TargetMode="External"`, whose `Target`
  (e.g. `测试.xlsx`) is where the relative/absolute path actually lives — the bracketed name in the
  formula is only an index, not the path.
- A workbook-level `<externalReferences>/<externalReference r:id="…"/>` entry plus rel wiring,
  assigning each external workbook a 1-based index.
- The cell formula then uses that **index** in brackets: `[1]Sheet1!A1`, not the filename. Writing the
  raw filename in brackets produces a formula Excel cannot bind — the reported "does nothing".

## Open questions

- **Public API shape:** a first-class cell value kind carrying `{path, sheet, cellOrRange,
  cachedValue}` (structured, in the spirit of a typed API) versus parsing `[name]Sheet!ref` magic
  strings (fragile; collides with defined-name and table syntax). Prefer the structured value.
- **Relative vs absolute paths and `TargetMode="External"`:** let the user control the relationship
  `Target` so relative sibling paths work; document that Excel resolves relative to the host file's
  folder (legacy fell back to the user's Documents folder).
- **Cached values:** write a plausible cached result so the reference displays before recalculation;
  policy when unknown.
- **Read side:** parse existing externalLink parts + external references on import, surface them as
  structured external-reference values, and round-trip faithfully (index remapping, rel-id
  uniqueness).
- **Scope:** whole external workbooks, named ranges in external workbooks, and DDE/OLE links are
  distinct sub-features; start with cell/range references to another `.xlsx` by path.

Related: `defined-name-formula-expression`, `internal-hyperlink-target-portability`,
`formula-recalculation-expectations`, `cross-sheet-reference-preserved-in-formula-and-validation`,
`formula-cell-value-type-minimal-required-fields`.
