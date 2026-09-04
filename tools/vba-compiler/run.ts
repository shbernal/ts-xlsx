#!/usr/bin/env node
// vba-compiler harness: orchestrator.
//
// Turns a declarative spec of VBA module source into a genuinely-compiled artifact by driving a real,
// headless Excel through the VBIDE object model (compile.ps1):
//   spec.json  ->  Excel VBIDE inject + compile + save  ->  vbaProject.bin (or a whole .xlsm)
//
// WHY THIS EXISTS: Excel does NOT recompile VBA from source on open. A module ships its compiled p-code
// and Excel runs that; a from-scratch or byte-spliced project with absent/mismatched p-code throws
// "Invalid data format" or silently runs stale code (recorded finding 2026-07-24). Genuinely compiled,
// source-matched p-code is a hard requirement, and only a real Excel can produce it. Hence this
// offline build tool. The shipped library stays pure-TS: it attaches the emitted bytes verbatim via
// `Workbook.vbaProjectBytes`.
//
// This is a PROBE/build tool, NOT a test. It is Windows/Excel-bound and never runs in CI; its output is
// a recorded artifact that seeds a committed corpus fixture (ADR 0012/0013 seed+lock split). It also
// needs Trust access to the VBA project object model
// (HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1). See README.md.
//
// Usage:  node tools/vba-compiler/run.ts <spec.json> --out <vbaProject.bin | out.xlsm>

import {readFileSync} from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

import {failWith, runPwsh} from '../pwsh.ts';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const COMPILE_PS1 = path.join(HERE, 'compile.ps1');
// Annotated, not inferred. TypeScript performs never-return control-flow analysis only for a
// function declaration or a const with an explicit type, so without this every `fail(...)` reads as
// an ordinary call and the code after it as reachable.
const fail: (message: string) => never = failWith('vba-compiler');
const PWSH_TIMEOUT_MS = 180_000;

/** A module to author: its VBA name, kind, and source (without a leading `Attribute VB_Name` line). */
interface ModuleSpec {
  readonly name: string;
  readonly kind: 'procedural' | 'class' | 'designer' | 'document';
  readonly source: string;
}

/** The compile spec: the modules to author, and, for editing an existing project, a base workbook. */
interface CompileSpec {
  readonly modules: readonly ModuleSpec[];
  /** Path to an existing `.xlsm` to edit in place (required for `document`/`designer` modules). */
  readonly base?: string;
}

/** The blob compile.ps1 emits on stdout. */
interface CompileResult {
  readonly ok: boolean;
  readonly mode: 'in-place' | 'from-scratch';
  readonly out: string;
  readonly modules: readonly {name: string; action: string; kind: string}[];
  readonly error: string | null;
}

// Refuse to run on a host without pwsh or a registered Excel COM server, so the tool never masquerades a
// missing dependency as a silent no-op.
async function assertExcelAvailable(): Promise<void> {
  const probe = await runPwsh(
    ['-Command', "if ([Type]::GetTypeFromProgID('Excel.Application')) { 'ok' } else { 'missing' }"],
    15_000,
  );
  if (probe.spawnError) {
    fail(
      'PowerShell (pwsh) was not found. The VBA compiler requires a Windows host with pwsh and Excel Desktop installed; it is not runnable here.',
    );
  }
  if (probe.stdout.trim() !== 'ok') {
    fail(
      'No registered Excel COM server (ProgID Excel.Application). The VBA compiler requires Excel Desktop installed on this Windows host.',
    );
  }
}

const VALID_KINDS = new Set(['procedural', 'class', 'designer', 'document']);

// `Array.isArray` is declared `arg is any[]`, and a `readonly T[]` is not assignable to `any[]`, so
// narrowing a readonly array through it widens every element to `any`, and the member reads in
// `readSpec` below would then be unchecked. This predicate does the same runtime test and keeps the
// declared element type.
function isReadonlyArray<T>(value: readonly T[] | undefined): value is readonly T[] {
  return Array.isArray(value);
}

function readSpec(specPath: string): CompileSpec {
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(specPath, 'utf8'));
  } catch (error) {
    fail(`could not read/parse spec ${specPath}: ${(error as Error).message}`);
  }
  const s = parsed as Partial<CompileSpec>;
  if (!isReadonlyArray(s.modules) || s.modules.length === 0) {
    fail(`spec ${specPath} must have a non-empty "modules" array`);
  }
  for (const m of s.modules) {
    if (!m || typeof m.name !== 'string' || m.name === '')
      fail('every module needs a non-empty "name"');
    if (!VALID_KINDS.has(m.kind))
      fail(`module '${m.name}': kind must be one of ${[...VALID_KINDS].join(', ')}`);
    if (typeof m.source !== 'string') fail(`module '${m.name}': "source" must be a string`);
    if ((m.kind === 'document' || m.kind === 'designer') && !s.base) {
      fail(`module '${m.name}': a '${m.kind}' module requires a "base" workbook to edit in place`);
    }
  }
  return s as CompileSpec;
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  const specPath = argv.find((a) => !a.startsWith('--'));
  if (!specPath)
    fail('usage: node tools/vba-compiler/run.ts <spec.json> --out <file.bin|file.xlsm>');
  const outIdx = argv.indexOf('--out');
  const outPath = outIdx >= 0 ? argv[outIdx + 1] : undefined;
  if (!outPath)
    fail('missing required --out <path> (a *.bin extracts vbaProject.bin, else a .xlsm)');

  await assertExcelAvailable();
  readSpec(specPath); // validate fail-closed before spawning Excel

  const run = await runPwsh(
    ['-File', COMPILE_PS1, '-Spec', path.resolve(specPath), '-Out', path.resolve(outPath)],
    PWSH_TIMEOUT_MS,
  );
  if (run.code !== 0 || run.stdout.trim() === '') {
    fail(`compile.ps1 failed (code ${run.code}): ${run.stderr.trim() || '(no stderr)'}`);
  }

  let result: CompileResult;
  try {
    result = JSON.parse(run.stdout) as CompileResult;
  } catch {
    fail(`compile.ps1 did not emit valid JSON:\n${run.stdout}`);
  }
  if (!result.ok) fail(`compilation failed: ${result.error ?? '(no error message)'}`);

  process.stderr.write(
    `vba-compiler: ${result.mode}, ${result.modules.map((m) => `${m.name}(${m.action})`).join(', ')} -> ${outPath}\n`,
  );
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
}

await main();
