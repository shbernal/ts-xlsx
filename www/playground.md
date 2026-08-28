---
title: Playground
description: Run ts-xlsx on a workbook in your own tab, and read every byte it produced.
editLink: false
aside: false
pageClass: page-wide
---

# Playground

<ClientOnly>
  <Playground />
</ClientOnly>

## What the boxes are

Each box is one function a consumer calls, and the time beside it is the time that function
took, measured around the call and nothing else.

- **`writeXlsx(workbook)`** turns a workbook into an `.xlsx` package. A sample starts here,
  because there is a workbook and no bytes yet.
- **`readXlsx(bytes)`** turns a package back into a workbook. A file you drop starts here,
  because there are bytes and no workbook.
- **The round trip** reads, writes, reads again, and compares the two reads. It compares the
  models rather than the bytes, because the bytes are allowed to differ: entry order and a
  shared-strings table are choices a writer may legally make differently from the one that
  produced your file. What must not differ is what a consumer sees. It also compares the two
  writes byte for byte, because this writer does promise that much.

A box that fails goes red and shows the library's own error message, and the boxes after it
say they did not run. That is deliberate. The failure taxonomy is a documented part of the
API, and a demo that folded it into "something went wrong" would be hiding the most useful
thing it had to show you.

## Nothing you drop leaves this tab

There is no upload here, and no server to upload to. The site is static files on GitHub
Pages, and the modules this page runs contain no `fetch`, no form and no url. Your file is
read with `FileReader`, handed to the library as a `Uint8Array`, and dropped when you pick
something else.

That is a property of the code rather than a promise about it. If you would rather check
than take it on trust, open your browser's network panel and drop a file: nothing is
requested.

## A preserved part is not a gap

Some parts of a real workbook are things this library does not model: a pivot cache, a
slicer, a linked-workbook reference. It reads their bytes, keeps them, and writes them back
untouched, so a file that goes through a read and a write does not quietly lose them. Those
parts are marked **preserved** in the package panel.

Seeing one is the system working. A library that modelled everything would be a library that
had to be finished before it could be safe; preserving what it does not model is what lets it
be safe first.

## What this does not prove

The built-in samples are written by this library and read back by this library. That is a
real check, and it is the weaker of the two available.

The stronger one is a file this library did not write. Drop one, and you are testing the
reader against a producer with its own habits, which is where a spreadsheet library actually
gets hard. Nothing on this page can tell you whether the workbook you get back opens cleanly
in Excel, because your browser cannot ask Excel. That question is answered elsewhere: the
regression corpus runs every case against real fixtures, and a separate gate validates
generated packages against Microsoft's own schema and semantic validator. This page is the
part of that you can watch happen.

The sheet is painted with the value's own text and the cell's font, fill and alignment. The
number format is shown when you hover a cell, and is deliberately not applied: this library
models formats and does not implement a formatting engine, so rendering `$1,234.50` here
would be showing you a cell it cannot actually produce.
