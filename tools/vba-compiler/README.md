# tools/vba-compiler

Turns VBA module **source** into a genuinely-compiled `vbaProject.bin` (or a whole `.xlsm`) by driving
a real, headless Excel through the VBIDE object model.

## Why this exists

Excel does **not** recompile VBA from source on open. Every module ships its compiled p-code
(PerformanceCache) inside `vbaProject.bin`, and Excel runs **that** — the source text is only recompiled
when a human opens the VBE and forces it. So a `vbaProject.bin` synthesized from source alone, or
byte-spliced with absent/mismatched p-code, does one of two bad things when macros are enabled:

- **throws `Invalid data format`** and deletes the project (no p-code, or malformed p-code), or
- **silently runs stale p-code** and ignores the new source entirely (mismatched p-code) — the worst
  case, because the file "opens clean".

(Recorded finding, 2026-07-24 — see `docs/decisions/0019-*`.) The only way to get genuinely compiled,
source-matched p-code is to let a real Excel compile it. This offline tool does exactly that; the shipped
`@shbernal/ts-xlsx` library stays **pure-TS with zero Office dependency** and simply attaches the
emitted bytes via `Workbook.vbaProjectBytes`.

## This is a PROBE / build tool — never CI

Same contract as `tools/excel-oracle`: Windows + a licensed Excel Desktop, inherently slow, never wired
into CI. Its output is a **recorded artifact that seeds a committed corpus fixture** (ADR 0012/0013
seed+lock split). The pure-TS attach round-trip is what runs in CI.

## Prerequisites

- Windows with Excel Desktop installed and COM-registered, plus `pwsh`.
- **Trust access to the VBA project object model** must be ON:
  `HKCU\Software\Microsoft\Office\<version>\Excel\Security\AccessVBOM = 1`
  (Excel ▸ File ▸ Options ▸ Trust Center ▸ Macro Settings ▸ "Trust access to the VBA project object
  model"). Off by default — VBIDE automation silently fails without it.

## Usage

```
node tools/vba-compiler/run.ts <spec.json> --out <vbaProject.bin | out.xlsm>
```

- `--out` ending in `.bin` → extracts and writes just `xl/vbaProject.bin` (attach it in pure TS).
- any other `--out` → writes the whole macro-enabled `.xlsm`.

### Spec format

```jsonc
{
  // Optional: an existing .xlsm to edit IN PLACE. Required for `document`/`designer` modules, and the
  // right mode for editing a real workbook's code-behind (ThisWorkbook/Sheet1) — the modules bind to
  // the real host. Omit for a from-scratch project.
  "base": "./path/to/workbook.xlsm",
  "modules": [
    { "name": "Module1", "kind": "procedural", "source": "Function AddThem(a, b)\r\n  AddThem = a + b\r\nEnd Function" },
    { "name": "Widget",  "kind": "class",      "source": "..." }
  ]
}
```

- `kind`: `procedural` | `class` | `designer` | `document`. A module whose `name` already exists in the
  `base` has its code **replaced**; a new name is **added**.
- **Do not** put a leading `Attribute VB_Name = "..."` line in `source` — the tool sets the name from
  `name`, and a duplicate silently breaks compilation. (The tool strips a leading `VB_Name` line
  defensively; keep other `VB_*` attributes for document modules.)

### Two modes

| Mode | When | Output |
|---|---|---|
| **from-scratch** (no `base`) | procedural/class modules for a fresh project | `.bin` to attach, or a `.xlsm` |
| **in-place** (with `base`) | edit an existing project, or author `document`/`designer` code-behind | the edited `.xlsm` |

**Prefer in-place for `document` code-behind.** A from-scratch `.bin` carries the throwaway workbook's
own empty `ThisWorkbook`/`Sheet1` document modules; attaching it to a different workbook works for
procedural/class macros but is not the way to author real code-behind — edit the target workbook in
place instead, so the code binds to its real sheets.

## Verifying output

A compiled artifact is only trustworthy once a **known macro from the authored source actually runs**.
Verify with the execute probe (never accept "opens clean" alone):

```
pwsh -File .claude/skills/excel-gui-automation/scripts/execute-verdict.ps1 `
    -Path out.xlsm -RunMacro AddThem -RunArgs 4,5 -ExpectResult 9
```

`passed: true` with the expected `runResult` confirms the p-code is real and matches the source.
