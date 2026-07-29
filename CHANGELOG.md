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

### Added

- **BREAKING: every deliberate failure now descends from one `XlsxError`, and carries a `code`.**
  `catch (e) { if (e instanceof XlsxError) … }` is the whole answer to "was that this library?" —
  previously five typed classes shared no ancestor, and the model's own validation threw bare `Error`
  distinguishable only by string-matching the message. `error.code` is the coarse branch
  (`'unsupported-format'` | `'malformed-input'` | `'authoring'` | `'internal'`); `error.name` and
  `instanceof` remain the exact one. See "How a failure is reported" in
  [docs/architecture.md](docs/architecture.md).

  New classes: `AuthoringError` (a document that cannot exist — 50 sites that used to throw bare
  `Error`), `PackageReadError` (a refused inflate, previously recognised by a message prefix),
  `XmlParseError` (malformed markup, previously a native `SyntaxError` that escaped `readXlsx`
  indistinguishable from a caller's own), `XlsxParseError`, and `InternalError` (an invariant of ours
  that did not hold). The five existing classes — `UnsupportedFormatError`, `XlsbParseError`,
  `VbaParseError`, `VbaAuthorError`, `CustomUiParseError` — keep their names, messages and fields and
  gain the ancestry.

  What did **not** change: scalar argument validation stays native `RangeError` / `SyntaxError` /
  `TypeError`, and every error message is byte-identical. The break is the *type* of a caught error,
  which matters if you catch `SyntaxError` around XML parsing or switch on `constructor`.

- **`Workbook.addTableStyle({name, elements})` — custom table styles are authorable.** A workbook can
  now define its own named table styles beside Excel's built-in gallery, and a table reaches one by
  putting that name in `TableStyleInfo.name`. Each element names a region (`wholeTable`, `headerRow`,
  `firstRowStripe`, … — all 28 of `ST_TableStyleType`) and carries a `DifferentialStyle`, interned into
  the same shared `<dxfs>` table conditional formatting uses, so two elements painted alike cost one
  entry. A stripe may set its band width with `size`. Authoring a name a source file already defined
  overrides that definition rather than adding an ambiguous second one.

  Verified against Excel Desktop, not just the schema: a table style is a cross-part correspondence
  every part of which can be valid while the table still renders unstyled. Excel registers the
  authored style in the workbook's gallery and paints from it — see
  `docs/knowledge/specs/custom-table-styles.md`.

  An unnamed style, or a `size` outside the four stripe types, throws — both otherwise produce a file
  Excel opens cleanly and then ignores. `TableStyleInfo.name` itself stays unvalidated, deliberately;
  the type documents why.

- **`Workbook.setTheme({colors, fonts})` — the workbook's palette is authorable.** A colour picked
  from a spreadsheet's theme row is written as `theme="4"`, a reference resolved at render time, so
  setting `accent1` restyles every cell, chart and table style that follows the theme at once — the
  only way to recolour a workbook without touching a cell. Any subset of the twelve colour-scheme
  slots and either of the two typefaces can be set, and calls merge.

  It generates *over* the existing theme rather than replacing it. The format scheme — the gradient,
  line and effect styles a designer authored — rides through untouched, a slot left unnamed keeps its
  source encoding (`dk1`/`lt1` stay `<a:sysClr>`, so they still follow the viewer's window colours),
  and a theme that references a picture keeps that relationship. `Workbook.themeColors` and
  `themeFonts` report the effective theme. A malformed colour throws at the setter, because Excel does
  not report one — it renders the slot as flat black.

- **`Workbook.resolveColor(color)` — a themed or indexed colour now resolves to a concrete ARGB.**
  A `Color` read from a file often carries no colour at all, only a reference: `{theme: 4}` into the
  workbook theme's scheme, or `{indexed: 2}` into the legacy 64-entry palette, either optionally with
  a `tint`. Resolution follows the workbook's *own* theme and its own custom `<indexedColors>` when it
  declares one, and applies the tint last. `Workbook.themeColors` exposes the scheme it resolves
  against. Two things it deliberately does not do: `indexed="64"`/`65` (the system foreground and
  background) resolve to `undefined` rather than to invented black and white, and nothing is written
  back into the model — the `Color` keeps the encoding its file used, so a round-trip still emits
  `theme="4" tint="0.4"` and the cell keeps its link to the theme.

  Note the index order: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is *not* the order the
  slots appear in the theme part. Verified against Excel Desktop; see
  `docs/knowledge/specs/theme-color-index-order.md`.

### Fixed

- **A workbook's theme part is no longer destroyed on round-trip.** Reading a file and writing it
  straight back replaced its theme with the default Office one, so every `theme="n"` colour and
  `scheme="major|minor"` font in the file silently re-rendered in the wrong brand — a branded
  workbook came back in Office blue. The source theme is now preserved verbatim, reached through the
  workbook's `.../theme` relationship rather than the conventional `xl/theme/theme1.xml` path, and
  re-emitted with the parts it references (a picture used as a themed fill) so its `r:embed` does not
  dangle. Exposed on the model as `Workbook.themePart` / `restoreThemePart`, opaque preserved XML in
  the same spirit as `restoreDifferentialStyles`.

- **Custom table styles and recent-colour swatches are no longer dropped on round-trip.** A workbook's
  `<tableStyles>` definitions and the default table/pivot styles it nominates were discarded when the
  stylesheet was regenerated, so a table asking for a custom style by name was left referencing
  nothing and rendered completely unstyled — a file that opens clean and looks wrong.
  `<colors><mruColors>` (the "Recent Colors" swatches) went the same way. Both are now preserved
  verbatim and exposed as `Workbook.tableStyles` / `restoreTableStyles` and `Workbook.mruColors` /
  `restoreMruColors`. Each `tableStyleElement`'s `dxfId` keeps resolving because the differential-style
  table is re-emitted at its original indices, and the namespace prefixes Excel stamps on a table style
  (`xr9:uid`) are re-declared on the stylesheet root rather than left dangling.

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
