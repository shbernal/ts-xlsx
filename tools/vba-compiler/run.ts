#!/usr/bin/env node
// vba-compiler harness — orchestrator.
//
// Turns a declarative spec of VBA module source into a genuinely-compiled artifact by driving a real,
// headless Excel through the VBIDE object model (compile.ps1):
//   spec.json  ->  Excel VBIDE inject + compile + save  ->  vbaProject.bin (or a whole .xlsm)
//
// WHY THIS EXISTS: Excel does NOT recompile VBA from source on open. A module ships its compiled p-code
// and Excel runs that; a from-scratch or byte-spliced project with absent/mismatched p-code throws
// "Invalid data format" or silently runs stale code (recorded finding 2026-07-24). Genuinely compiled,
// source-matched p-code is a hard requirement, and only a real Excel can produce it — hence this
// offline build tool. The shipped library stays pure-TS: it attaches the emitted bytes verbatim via
// `Workbook.vbaProjectBytes`.
//
// This is a PROBE/build tool, NOT a test. It is Windows/Excel-bound and never runs in CI; its output is
// a recorded artifact that seeds a committed corpus fixture (ADR 0012/0013 seed+lock split). It also
// needs Trust access to the VBA project object model
// (HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1) — see README.md.
//
// Usage:  node tools/vba-compiler/run.ts <spec.json> --out <vbaProject.bin | out.xlsm>

import {spawn} from 'node:child_process';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const COMPILE_PS1 = path.join(HERE, 'compile.ps1');
const PWSH_TIMEOUT_MS = 180_000;
const MAX_OUTPUT_BYTES = 8 * 1024 * 1024;

/** A module to author: its VBA name, kind, and source (without a leading `Attribute VB_Name` line). */
interface ModuleSpec {
  readonly name: string;
  readonly kind: 'procedural' | 'class' | 'designer' | 'document';
  readonly source: string;
}

/** The compile spec: the modules to author, and — for editing an existing project — a base workbook. */
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

interface PwshResult {
  readonly code: number | null;
  readonly stdout: string;
  readonly stderr: string;
  readonly spawnError?: Error;
}

function fail(message: string): never {
  process.stderr.write(`vba-compiler: ${message}\n`);
  process.exit(1);
}

function runPwsh(args: readonly string[], timeoutMs: number): Promise<PwshResult> {
  return new Promise<PwshResult>((resolve) => {
    const child = spawn('pwsh', ['-NoProfile', '-NonInteractive', ...args], {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    let bytes = 0;
    let settled = false;
    const done = (r: PwshResult) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(r);
    };
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      done({code: null, stdout: Buffer.concat(stdout).toString('utf8'), stderr: 'timed out'});
    }, timeoutMs);
    const collect = (target: Buffer[]) => (chunk: Buffer) => {
      bytes += chunk.length;
      if (bytes > MAX_OUTPUT_BYTES) {
        child.kill('SIGKILL');
        done({
          code: null,
          stdout: Buffer.concat(stdout).toString('utf8'),
          stderr: 'output too large',
        });
        return;
      }
      target.push(chunk);
    };
    child.stdout.on('data', collect(stdout));
    child.stderr.on('data', collect(stderr));
    child.on('error', (spawnError) => done({code: null, stdout: '', stderr: '', spawnError}));
    child.on('close', (code) =>
      done({
        code,
        stdout: Buffer.concat(stdout).toString('utf8'),
        stderr: Buffer.concat(stderr).toString('utf8'),
      }),
    );
  });
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

function readSpec(specPath: string): CompileSpec {
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(specPath, 'utf8'));
  } catch (error) {
    fail(`could not read/parse spec ${specPath}: ${(error as Error).message}`);
  }
  const s = parsed as Partial<CompileSpec>;
  if (!Array.isArray(s.modules) || s.modules.length === 0) {
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
    `vba-compiler: ${result.mode} — ${result.modules.map((m) => `${m.name}(${m.action})`).join(', ')} -> ${outPath}\n`,
  );
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
}

await main();
