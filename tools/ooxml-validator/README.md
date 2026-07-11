# OOXML validator

This development-only console tool validates `.xlsx` packages with Microsoft's
`DocumentFormat.OpenXml.Validation.OpenXmlValidator`. It is an independent schema and semantic
oracle for generated workbooks; it is not part of the published JavaScript package and does not
replace the regression corpus or real Excel interoperability testing.

The project targets .NET 10 and pins `DocumentFormat.OpenXml` through both the project file and a
NuGet lockfile. Validation defaults to `FileFormatVersions.Microsoft365`.

```bash
npm run validate:ooxml -- workbook.xlsx another.xlsx
npm run test:ooxml
```

The command writes a deterministic JSON report to stdout. Exit code `0` means every input is clean,
`1` means validation or package-open errors were found, and `2` means the tool could not run because
of invalid arguments or an internal failure.
