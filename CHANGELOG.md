# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project adheres to
[Semantic Versioning](https://semver.org/) — see
[ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md) for the
versioning policy, including why the first published version is `1.0.0` rather than a
`0.x` series.

This file tracks changes from its introduction forward. Earlier history — the full
ExcelJS-to-`ts-xlsx` rewrite — is recorded in `git log` and the [ADR series](docs/decisions/).

## [Unreleased]

### Removed

- **BREAKING: the pure-TS VBA source-authoring API is gone.** `writeVbaProject`,
  `Workbook.setVbaProject`, `editVbaModuleSources`, `Workbook.setVbaModuleSource`,
  `editXlsxVbaModuleSource`, `editXlsxVbaModuleSources`, `addVbaModule`, `Workbook.addVbaModule`, and
  `editXlsxVbaAddModule` are removed, along with the `VbaModuleSource`/`VbaProjectSpec` types. They
  emitted modules with no compiled p-code on the theory that Excel recompiles from source on open —
  which it does not: such files either fail to load ("Invalid data format") or silently run stale code.
  See [ADR 0019](docs/decisions/0019-vba-authoring-needs-real-pcode-recompile-cookie-retracted.md).

### Added

- **`tools/vba-compiler`** — an offline build tool that produces genuinely compiled, source-matched VBA
  p-code by driving a real headless Excel (VBIDE). Emits a `vbaProject.bin` (attach via
  `Workbook.vbaProjectBytes`) or a whole edited `.xlsm`. Windows + licensed Excel only; never in CI.

### Changed

- `removeVbaModule` and `addVbaReference` (and their `Workbook`/`editXlsxVba*` wrappers) no longer reset
  the `_VBA_PROJECT` stream — that reset crashed the VBA load on a project with real p-code. They now
  leave it byte-for-byte untouched; the `dir` stream carries the structural change. These edits are
  retained precisely because they never touch a module's compiled p-code.

- **Contributor tooling: the gate set has one name.** `node scripts/verify.ts` runs every gate
  concurrently and is what `pnpm test`, lefthook's `pre-push` and the turn-boundary hook all
  invoke; `--quick` is the inner loop and `--cached` exits immediately when the working tree is
  byte-for-byte the one that last passed. The corpus runner gained `--case`/`--json` and now
  prints only what needs attention (`--verbose` for the old listing), and the OOXML validator
  builds on demand instead of paying `dotnet run`'s project re-evaluation per call. Nothing in
  the published package changes. See
  [ADR 0022](docs/decisions/0022-verification-is-one-cached-parallel-entrypoint.md).
