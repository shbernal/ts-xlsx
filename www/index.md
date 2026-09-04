---
title: ts-xlsx
description: A TypeScript-first library for reading and writing xlsx spreadsheets. Synchronous, Uint8Array-native, one runtime dependency.
layout: page
pageClass: page-home
editLink: false
aside: false
---

<div class="home-hero xlsx-lattice vp-doc">

# Spreadsheets, read and written honestly

A TypeScript-first library for `.xlsx` and CSV. The buffered path is synchronous and speaks
`Uint8Array`, so writing a workbook is a function call that hands you bytes, and reading one
is a function call that hands you a model. One runtime dependency, and every behaviour
pinned by a regression case.

```shell
npm install @shbernal/ts-xlsx
```

<QuickStart />

<div class="home-doors">
  <Door href="/docs/guide/" title="Read the guide">Eight pages, from install to the browser.</Door>
  <Door href="/playground" title="Run it in this tab">Write, read and round-trip, with the bytes on screen.</Door>
  <Door href="/docs/api/" title="Browse the API">Generated from the public types.</Door>
</div>

</div>

<div class="home-body vp-doc">

<p class="xlsx-kicker">The dependency tree</p>

## {{ $facts.runtimeDependencies }} runtime dependency

<code>{{ $facts.runtimeDependencyNames }}</code>, for zip. Everything else is written here: the XML
reader, the OOXML model, the CSV codec, the compound-file reader that opens a macro project.
A large part of why this library exists is that the project it forked from had a rotting
transitive tree, so the audit is part of CI and is expected to stay green.

<p class="xlsx-kicker">Types</p>

## The types are the contract

`strict`, plus `noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`. The
[API reference](/docs/api/) is generated straight from the published declarations, so it
cannot describe a shape the compiler would reject, and a wrong page in it is a wrong JSDoc
comment in the source rather than a documentation task
([ADR-0006](/docs/decisions/0006-docs-from-types)).

A cell's value is one precisely typed union rather than a `type` field you set alongside a
value. What you assign is what the cell is.

<p class="xlsx-kicker">Correctness</p>

## {{ $facts.corpusCases }} regression cases, harvested from a real backlog

This library is a hard fork of ExcelJS, and the fork's first job was to get the knowledge out
before discarding the code. Every credible bug, reproduction and edge case in that backlog
became a corpus case written against _behaviour_ rather than against an implementation, which
is how they outlived the rewrite they were built to survive. A bug without a case is a bug
that will return.

Beside them sit {{ $facts.specNotes }} hand-authored spec notes, the evidence a case is
written from, and {{ $facts.decisionRecords }} decision records saying why each fork in the
road went the way it did, including the ones that were later retracted.

<p class="xlsx-kicker">Output</p>

## Emitted files are validated, not assumed

"It opens in Excel" is not a test. Generated packages are checked against Microsoft's own
`OpenXmlValidator`, schema and semantics both, as an independent oracle rather than as this
library grading its own homework
([ADR-0002](/docs/decisions/0002-ooxml-validation-oracle)).

The bytes are also a pure function of the model: an unchanged workbook written twice produces
two identical archives, because entry timestamps are pinned rather than clocked
([ADR-0032](/docs/decisions/0032-package-output-is-reproducible)). A committed `.xlsx`
therefore changes only when something about it changed.

<p class="xlsx-kicker">The honest part</p>

## Three states, and there is no fourth

Every part of a workbook this library meets is in exactly one of them.

<div class="home-states">
  <div class="home-state">
    <h3>Modelled</h3>
    <p>The reader understands it and the writer can rebuild it. Cells, styles, formulas,
    tables, merges, panes, validation, conditional formats, comments, images, page setup.</p>
  </div>
  <div class="home-state">
    <h3>Preserved</h3>
    <p>The model does not interpret it, and the bytes cross a load and save untouched. Pivot
    caches, slicers, charts, shapes, linked-workbook references, a VBA project.</p>
  </div>
  <div class="home-state">
    <h3>Refused</h3>
    <p>The library says so, with a typed error naming which of four kinds of failure it was,
    rather than guessing and handing back something plausible.</p>
  </div>
</div>

There is deliberately no "approximated" state, and that absence is the point: a library that
silently half-understands a part is one whose output you cannot trust without opening it.

Say the limit in the same breath as the claim. Charts, vector shapes, slicers and legacy form
controls are in the **preserved** column, not the modelled one: a workbook that has them
keeps them, and there is no API to author a new one
([ADR-0014](/docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0)). If that
is what you came for, you want a different library, and it is better that you learn it here
than after adopting this one.

<p class="xlsx-kicker">Untrusted input</p>

## Every parser path assumes the file is hostile

Reading a spreadsheet means running someone else's bytes through your process. Inflation is
bounded by counting the output actually produced, never by trusting the archive's declared
sizes, so a zip bomb that lies about its size is refused all the same. XML entities are
decoded but never expanded. Neither defence is configurable, and neither can be switched off.

<p class="xlsx-kicker">See for yourself</p>

## The playground is the claim, running

The [playground](/playground) writes a workbook, reads it back, round-trips it, and shows you
the emitted package part by part, in your own tab, with no server behind it. Drop a
spreadsheet of your own and you are testing the reader against a producer with its own
habits, which is where a spreadsheet library actually gets hard.

<div class="home-doors">
  <Door href="/docs/guide/" title="Guide">{{ $facts.guidePages }} pages, in reading order.</Door>
  <Door href="/docs/api/" title="API reference">Generated from the types.</Door>
  <Door href="/docs/migrating-from-exceljs" title="Coming from ExcelJS?">What changed, and why it changed.</Door>
  <Door href="https://github.com/shbernal/ts-xlsx" title="The repository">Source, records and the corpus.</Door>
</div>

</div>
