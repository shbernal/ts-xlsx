# OOXML validator

This development-only console tool validates `.xlsx` packages with Microsoft's
`DocumentFormat.OpenXml.Validation.OpenXmlValidator`. It is an independent schema and semantic
oracle for generated workbooks; it is not part of the published JavaScript package and does not
replace the regression corpus or real Excel interoperability testing.

The project targets .NET 10 and pins `DocumentFormat.OpenXml` through both the project file and a
NuGet lockfile. Validation defaults to `FileFormatVersions.Microsoft365`.

```bash
node scripts/ooxml-validator.ts workbook.xlsx another.xlsx   # or: pnpm run validate:ooxml -- …
node test/ooxml-validation/run.ts                            # or: pnpm run test:ooxml
```

Both entry points go through `scripts/ooxml-validator.ts`, which builds the assembly only when it is
older than the sources beside it and then invokes it directly — `dotnet run --project` re-evaluates
the project on every call and costs more than the validation itself. That build restores in locked
mode, so a dependency the lockfile does not pin still fails loudly; it is simply not re-checked on
runs where nothing changed. Pass several files in one call: process startup dominates per-file cost.

The command writes a deterministic JSON report to stdout. Exit code `0` means every input is clean,
`1` means validation or package-open errors were found, and `2` means the tool could not run because
of invalid arguments or an internal failure.
