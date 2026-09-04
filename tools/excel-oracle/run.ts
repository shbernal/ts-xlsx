#!/usr/bin/env node
// Excel-oracle harness: orchestrator.
//
// One command turns a declarative probe into a recorded Excel-Desktop observation:
//   emit (spec -> .xlsx)  ->  observe (headless COM readback + re-save)  ->  collect (parse + canonical
//   ref readback in Node).
//
// This tool is a PROBE, not a test. It is Windows/Excel-bound and never runs in CI; the corpus runner
// (`node test/corpus/run.ts`) must never depend on it. Its output is a *recorded fact* that seeds a
// case; a Tier-2 seam fact is what locks that case and runs in CI (ADR 0012, seed+lock split).
//
// Usage:  node tools/excel-oracle/run.ts <probe.json> [--out <observation.json>] [--keep]
//
// Self-guards: it refuses to run (loud, non-zero exit) if pwsh or a registered Excel COM server is
// absent, so on a non-Excel host it degrades with a clear message rather than silently emitting empty
// facts.

import {readFileSync, rmSync, writeFileSync} from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

import {strFromU8, unzipSync} from 'fflate';

import {scratchDir} from '../../scripts/repo.ts';
import {failWith, runPwsh} from '../pwsh.ts';
import {emitProbe, type ProbeSpec} from './emit-probe.ts';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const OBSERVE_PS1 = path.join(HERE, 'observe.ps1');
// Annotated, not inferred. TypeScript performs never-return control-flow analysis only for a
// function declaration or a const with an explicit type, so without this every `fail(...)` reads as
// an ordinary call and the code after it as reachable.
const fail: (message: string) => never = failWith('excel-oracle');
const PWSH_TIMEOUT_MS = 120_000;

/** A probe file: the workbook to emit, and what to observe once Excel has opened it. */
interface Probe {
  readonly invariant: string;
  readonly description?: string;
  readonly spec: ProbeSpec;
  readonly observe: {readonly cells: readonly string[]; readonly resave?: boolean};
  /** The authored interpretation the probe records; echoed verbatim into the observation sidecar. */
  readonly verdict?: string;
}

/** One cell as Excel Desktop reported it back over COM. */
interface CellReadback {
  readonly address: string;
  readonly hasFormula: boolean;
  readonly formula: string;
  readonly value: string;
}

/** The raw observation blob observe.ps1 emits on stdout. */
interface RawObservation {
  readonly version: string | null;
  readonly build: number | null;
  readonly openThrew: boolean;
  readonly openError: string | null;
  readonly repaired: boolean | null;
  readonly workbookName: string | null;
  readonly cells: readonly CellReadback[];
  readonly resaved: boolean;
  readonly resavedPath: string | null;
  readonly resaveThrew: boolean;
  readonly resaveError: string | null;
}

// Refuse to run on a host without pwsh or a registered Excel COM server, so the harness never
// masquerades a missing dependency as an empty observation.
async function assertExcelAvailable(): Promise<void> {
  const probe = await runPwsh(
    ['-Command', "if ([Type]::GetTypeFromProgID('Excel.Application')) { 'ok' } else { 'missing' }"],
    15_000,
  );
  if (probe.spawnError) {
    fail(
      'PowerShell (pwsh) was not found. The Excel oracle requires a Windows host with pwsh and Excel Desktop installed; it is not runnable here.',
    );
  }
  if (probe.stdout.trim() !== 'ok') {
    fail(
      'No registered Excel COM server (ProgID Excel.Application). The Excel oracle requires Excel Desktop installed on this Windows host.',
    );
  }
}

function readProbe(probePath: string): Probe {
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(probePath, 'utf8'));
  } catch (error) {
    fail(`could not read/parse probe ${probePath}: ${(error as Error).message}`);
  }
  const p = parsed as Partial<Probe>;
  if (typeof p.invariant !== 'string' || p.invariant.length === 0) {
    fail(`probe ${probePath} is missing a non-empty "invariant"`);
  }
  if (!p.spec || !Array.isArray(p.spec.sheets)) {
    fail(`probe ${probePath} is missing "spec.sheets"`);
  }
  if (!p.observe || !Array.isArray(p.observe.cells)) {
    fail(`probe ${probePath} is missing "observe.cells"`);
  }
  return p as Probe;
}

/** Read back the shared-formula `<f>` elements Excel itself wrote: its canonical form for the group. */
function canonicalSharedFormulas(resavedPath: string): Record<string, string[]> {
  const zip = unzipSync(new Uint8Array(readFileSync(resavedPath)));
  const out: Record<string, string[]> = {};
  for (const name of Object.keys(zip)) {
    const match = /^xl\/worksheets\/(sheet\d+)\.xml$/.exec(name);
    const part = zip[name];
    if (!match || !part) continue;
    const xml = strFromU8(part);
    const els =
      xml.match(/<f\b[^>]*\bt="shared"[^>]*>[\s\S]*?<\/f>|<f\b[^>]*\bt="shared"[^>]*\/>/g) ?? [];
    if (els.length > 0) out[match[1] as string] = els;
  }
  return out;
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  const probePath = argv.find((a) => !a.startsWith('--'));
  if (!probePath)
    fail('usage: node tools/excel-oracle/run.ts <probe.json> [--out <file>] [--keep]');
  const outIdx = argv.indexOf('--out');
  const outPath = outIdx >= 0 ? argv[outIdx + 1] : undefined;
  const keep = argv.includes('--keep');

  await assertExcelAvailable();
  const probe = readProbe(probePath);
  const resave = probe.observe.resave !== false;

  const work = scratchDir('excel-oracle');
  const xlsxPath = path.join(work, `${probe.invariant}.xlsx`);
  const resavedPath = path.join(work, `${probe.invariant}.excel-resaved.xlsx`);

  try {
    emitProbe(probe.spec, xlsxPath);

    const args = ['-File', OBSERVE_PS1, '-Path', xlsxPath, '-Cells', probe.observe.cells.join(',')];
    if (resave) args.push('-SaveAsPath', resavedPath);
    else args.push('-NoResave');

    const run = await runPwsh(args, PWSH_TIMEOUT_MS);
    if (run.code !== 0 || run.stdout.trim() === '') {
      fail(`observe.ps1 failed (code ${run.code}): ${run.stderr.trim() || '(no stderr)'}`);
    }

    let raw: RawObservation;
    try {
      raw = JSON.parse(run.stdout) as RawObservation;
    } catch {
      fail(`observe.ps1 did not emit valid JSON:\n${run.stdout}`);
    }

    const canonicalSharedFormulasBySheet =
      raw.resaved && !raw.resaveThrew && raw.resavedPath
        ? canonicalSharedFormulas(raw.resavedPath)
        : {};

    const observation = {
      invariant: probe.invariant,
      ...(probe.description !== undefined ? {description: probe.description} : {}),
      // The probe that produced this observation, so the recorded fact points back at its own inputs.
      probeSpecRef: path.relative(process.cwd(), probePath).replaceAll('\\', '/'),
      excel: {version: raw.version, build: raw.build},
      capturedAt: new Date().toISOString().slice(0, 10),
      openClass: 'automation-open (DisplayAlerts=false, AutomationSecurity=ForceDisable)',
      open: {threw: raw.openThrew, error: raw.openError, repairedMarker: raw.repaired},
      cells: raw.cells,
      resave: {
        attempted: raw.resaved,
        threw: raw.resaveThrew,
        error: raw.resaveError,
        canonicalSharedFormulasBySheet,
      },
      // The authored interpretation this probe records: the conclusion the observation justifies.
      // Echoed from the probe (not derived) so the seeding case's provenance points at fact + verdict
      // together; a Tier-2 seam fact is still what locks the behavior in CI (ADR 0012).
      ...(probe.verdict !== undefined ? {verdict: probe.verdict} : {}),
    };

    const json = JSON.stringify(observation, null, 2);
    if (outPath) {
      writeFileSync(outPath, `${json}\n`);
      process.stderr.write(`excel-oracle: wrote observation to ${outPath}\n`);
    }
    process.stdout.write(`${json}\n`);
  } finally {
    if (!keep) rmSync(work, {recursive: true, force: true});
    else process.stderr.write(`excel-oracle: kept working files under ${work}\n`);
  }
}

await main();
