# Text metrics

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `estimateWrappedLines`

<sub>function</sub>

The number of lines `text` occupies when wrapped at `width` character units - the width unit a
column states, so `estimateWrappedLines(cell, sheet.getColumn(2).width ?? 8.43)` is the shape of
the call.

A hard break opens a line of its own and what follows wraps independently, matching how Excel
lays a wrapped cell out. The empty string is one line, not zero: a cell always occupies its row.

An estimate, and only ever that. It counts characters, so it is exact for a monospaced face and
approximate for every other - a run of `W`s wraps sooner on screen than this predicts, a run of
`i`s later. Excel's own auto-fit measures glyphs; this exists so that a writer can state *a*
height rather than leave the sheet to be laid out lazily on open, and being within a line of the
truth is what that needs.

```ts
function estimateWrappedLines(text: string, width: number): number;
```

**Throws** — `RangeError` if `width` is not a positive finite number - a column of zero width wraps
nothing, and silently answering `Infinity` or `NaN` would put that straight into a row height.
