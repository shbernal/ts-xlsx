#!/usr/bin/env node
// The OOXML oracle's entrypoint: build the .NET validator when it is stale, then invoke
// the built assembly directly.
//
// `dotnet run --project` looks correct and costs ~2.1 s warm (~6 s cold) because it
// re-evaluates the project and its dependency graph on every call, for a tool whose
// actual work is ~0.75 s. That overhead is paid per validated file-set, in the inner
// loop of exactly the workflow we want agents to reach for. So: decide staleness here,
// where it is one mtime comparison, and hand `dotnet` a path to the assembly.
//
// Usage (same contract as the tool itself — see tools/ooxml-validator/README.md):
//   node scripts/ooxml-validator.ts [--format Microsoft365] <file.xlsx> [more.xlsx ...]

import {spawn} from 'node:child_process';
import {readdir, stat} from 'node:fs/promises';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const PROJECT_DIR = join(ROOT, 'tools', 'ooxml-validator');
const PROJECT = join(PROJECT_DIR, 'OoxmlValidator.csproj');
const DLL = join(PROJECT_DIR, 'bin', 'Release', 'net10.0', 'OoxmlValidator.dll');

/** The tool's own "could not run" code; reused for failures reaching it. */
const TOOL_FAILURE = 2;

/**
 * What the assembly is built from. Everything else under the project directory is `bin/`
 * and `obj/` output, so a non-recursive read is exactly the source set — including
 * `packages.lock.json`, which is why a dependency bump still forces the locked-mode
 * restore that `dotnet build` performs below.
 */
const SOURCE_SUFFIXES = ['.cs', '.csproj', '.json'];

async function newestSourceMtime(): Promise<number> {
  const entries = await readdir(PROJECT_DIR, {withFileTypes: true});
  const sources = entries
    .filter((entry) => entry.isFile() && SOURCE_SUFFIXES.some((ext) => entry.name.endsWith(ext)))
    .map((entry) => join(PROJECT_DIR, entry.name));
  const times = await Promise.all(sources.map(async (file) => (await stat(file)).mtimeMs));
  return Math.max(0, ...times);
}

async function mtimeOrMissing(file: string): Promise<number | undefined> {
  try {
    return (await stat(file)).mtimeMs;
  } catch {
    return undefined;
  }
}

interface DotnetRun {
  readonly code: number;
  readonly captured: string;
}

/**
 * `stdout: 'capture'` buffers the child's stdout instead of passing it through — the build
 * needs it, because MSBuild reports its errors (a lock-file mismatch, a compile failure) on
 * *stdout*, and this process' stdout is reserved for the validator's JSON.
 */
function spawnDotnet(args: string[], stdout: 'inherit' | 'capture'): Promise<DotnetRun> {
  return new Promise((settle, fail) => {
    const child = spawn('dotnet', args, {
      cwd: ROOT,
      stdio: ['ignore', stdout === 'capture' ? 'pipe' : 'inherit', 'inherit'],
    });
    const chunks: Buffer[] = [];
    child.stdout?.on('data', (chunk: Buffer) => chunks.push(chunk));
    child.on('error', (error) =>
      fail(
        new Error(
          `could not run \`dotnet\`: ${error.message}\n` +
            'The OOXML oracle needs .NET 10 on PATH — see docs/agent-correctness-playbook.md.',
        ),
      ),
    );
    child.on('close', (code) =>
      settle({code: code ?? TOOL_FAILURE, captured: Buffer.concat(chunks).toString('utf8')}),
    );
  });
}

/**
 * Path to the built validator, building it first if it is missing or older than its
 * sources. Exported so the baseline harness (`test/ooxml-validation/run.ts`) resolves the
 * assembly the same way rather than keeping a second, drifting copy of this decision.
 */
export async function resolveValidator(): Promise<string> {
  const [built, sources] = await Promise.all([mtimeOrMissing(DLL), newestSourceMtime()]);
  if (built !== undefined && built >= sources) return DLL;

  // RestoreLockedMode makes the implicit restore fail on any dependency the lock file does
  // not already pin — the supply-chain check that used to live in a separate `dotnet
  // restore --locked-mode` invocation, now paid only when something actually changed.
  const {code, captured} = await spawnDotnet(
    ['build', PROJECT, '--configuration', 'Release', '--nologo', '-p:RestoreLockedMode=true'],
    'capture',
  );
  if (code !== 0) {
    throw new Error(`building the OOXML validator failed (dotnet build exit ${code})\n${captured}`);
  }
  return DLL;
}

if (import.meta.main) {
  try {
    const dll = await resolveValidator();
    // `pnpm run validate:ooxml -- file.xlsx` forwards the separator itself, and the tool
    // would reject it as a filename. Both spellings are documented; neither should surprise.
    const argv = process.argv.slice(2);
    if (argv[0] === '--') argv.shift();
    process.exitCode = (await spawnDotnet([dll, ...argv], 'inherit')).code;
  } catch (error: unknown) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = TOOL_FAILURE;
  }
}
