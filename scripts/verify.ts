#!/usr/bin/env node
// One entrypoint for the verification gates, run concurrently.
//
// The gates are independent processes with no shared state, so running them in
// sequence only buys the illusion of tidiness: it costs the sum of their times
// instead of the maximum. Assembling the chain by hand also lets gates go missing —
// `docs:check` and `constitution:check` are the two that get forgotten, because they
// are the two that rarely fail.
//
// Usage:
//   node scripts/verify.ts [--full | --quick] [--jobs <n>] [--list] [--cached]
//
//   --full        every gate — the same set lefthook runs pre-push and CI enforces.
//                 The default, because a bare invocation must never quietly skip the spine.
//   --quick       the inner loop — types, unit tests, and lint scoped to the files you
//                 have actually touched. No corpus, so it is not a substitute for --full.
//   --jobs <n>    how many gates to run at once (default 2; see runPool).
//   --list        print this mode's gate names and exit, without running them.
//   --cached      exit 0 immediately if this exact tree already passed this mode
//                 (see treeKey); otherwise run, and record the pass on success.
//
// Call this directly with `node`, not through `pnpm run`: the package-manager wrapper
// costs ~0.84 s per invocation, which is a tenth of the whole --quick budget.

import {spawn} from 'node:child_process';
import {createHash} from 'node:crypto';
import {mkdir, readFile, writeFile} from 'node:fs/promises';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const STAMP = join(ROOT, '.tmp', 'verify-stamp.json');

// Tools are invoked as node scripts against their package entrypoints rather than via
// node_modules/.bin, because a .bin entry on Windows is a .cmd shim that would force
// `shell: true` — and with it quoting rules that differ per platform.
const NODE = process.execPath;
const BIOME = resolve(ROOT, 'node_modules/@biomejs/biome/bin/biome');
const TSC = resolve(ROOT, 'node_modules/typescript/bin/tsc');

/** The directories `lint` covers; must stay in step with the `lint` package script. */
const LINT_ROOTS = ['src', 'scripts', 'test', 'tools'];
// CLAUDE.md §2 admits no warnings, but Biome exits 0 on them: without this a rule demoted to a
// warning (most of `style`, including noNonNullAssertion) is enforced by nothing at all.
const LINT_STRICT = '--error-on-warnings';
const LINTABLE = /\.(?:ts|js|mjs|cjs|json|jsonc)$/;
// Past this many changed files, an explicit list stops being cheaper than a whole-tree
// pass and starts crowding the OS argument limit. A codemod pays the 2 s.
const SCOPED_LINT_LIMIT = 100;

// How many gates run at once. Deliberately not derived from the core count — see
// runPool. Override with --jobs on a machine whose I/O is not the ceiling.
const DEFAULT_JOBS = 2;

type Mode = 'quick' | 'full';

interface Step {
  command: string;
  args: string[];
}

interface Gate {
  name: string;
  /** Run in order; the gate fails on the first step that does. */
  steps: Step[];
  /** Why the gate had nothing to do — set instead of steps, and reported as skipped. */
  vacuous?: string;
}

interface Result {
  gate: Gate;
  ok: boolean;
  ms: number;
  exit: number | null;
  output: string;
}

/** A bad invocation, not a failing gate: one legible line, no stack. */
class UsageError extends Error {}

interface Args {
  mode: Mode;
  list: boolean;
  jobs: number;
  cached: boolean;
}

function parseArgs(argv: string[]): Args {
  const args: Args = {mode: 'full', list: false, jobs: DEFAULT_JOBS, cached: false};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--full') args.mode = 'full';
    else if (arg === '--quick') args.mode = 'quick';
    else if (arg === '--list') args.list = true;
    else if (arg === '--cached') args.cached = true;
    else if (arg === '--jobs') {
      const value = Number(argv[++i]);
      if (!Number.isInteger(value) || value < 1) {
        throw new UsageError(`--jobs needs a positive integer, got ${argv[i]}`);
      }
      args.jobs = value;
    } else {
      throw new UsageError(
        `unrecognized argument: ${arg} (expected --full, --quick, --jobs, --list, --cached)`,
      );
    }
  }
  return args;
}

function run(step: Step): Promise<{exit: number | null; output: string}> {
  return new Promise((settle) => {
    const child = spawn(step.command, step.args, {cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe']});
    // One buffer for both streams: a reader wants the tool's diagnostics in the order
    // the tool emitted them, not stdout and stderr sorted into separate piles.
    const chunks: Buffer[] = [];
    child.stdout?.on('data', (chunk: Buffer) => chunks.push(chunk));
    child.stderr?.on('data', (chunk: Buffer) => chunks.push(chunk));
    child.on('error', (err) => settle({exit: null, output: `${step.command}: ${err.message}`}));
    child.on('close', (exit) => settle({exit, output: Buffer.concat(chunks).toString('utf8')}));
  });
}

async function capture(command: string, args: string[]): Promise<string> {
  const {output} = await run({command, args});
  return output;
}

/**
 * The lintable files that differ from HEAD — unstaged, staged and untracked. Biome's
 * own `--changed --since` was measured slower than a whole-tree pass (the diff it runs
 * internally eats the saving), so we compute the list once and hand it over explicitly.
 */
async function changedLintTargets(): Promise<string[]> {
  const lists = await Promise.all([
    capture('git', ['diff', '--name-only', '--diff-filter=ACMR']),
    capture('git', ['diff', '--cached', '--name-only', '--diff-filter=ACMR']),
    capture('git', ['ls-files', '--others', '--exclude-standard']),
  ]);
  const paths = new Set(
    lists
      .flatMap((list) => list.split('\n'))
      .map((line) => line.trim())
      .filter(Boolean),
  );
  return [...paths]
    .filter((path) => LINTABLE.test(path) && LINT_ROOTS.some((root) => path.startsWith(`${root}/`)))
    .sort();
}

function wholeTreeLint(): Gate {
  return {
    name: 'lint',
    steps: [{command: NODE, args: [BIOME, 'check', LINT_STRICT, ...LINT_ROOTS]}],
  };
}

function scopedLint(changed: string[]): Gate {
  if (changed.length === 0) return {name: 'lint', steps: [], vacuous: 'no changed files'};
  if (changed.length > SCOPED_LINT_LIMIT) return wholeTreeLint();
  return {
    name: 'lint',
    steps: [
      {
        command: NODE,
        // A changed path can be anything git reports; let Biome pass over what it does
        // not handle rather than failing the gate on an unmatched argument.
        args: [
          BIOME,
          'check',
          LINT_STRICT,
          '--no-errors-on-unmatched',
          '--files-ignore-unknown=true',
          ...changed,
        ],
      },
    ],
  };
}

/**
 * Declaration order is *start* order, longest pole first. The scheduler below runs a
 * bounded number of gates at a time, so the slowest gate must be in flight from the
 * outset or it becomes a tail nobody is overlapping with.
 */
async function gateSet(mode: Mode): Promise<Gate[]> {
  const lint = mode === 'full' ? wholeTreeLint() : scopedLint(await changedLintTargets());
  const gates: Gate[] = [];
  if (mode === 'full')
    gates.push({name: 'corpus', steps: [{command: NODE, args: ['test/corpus/run.ts']}]});
  gates.push(
    lint,
    {name: 'test:src', steps: [{command: NODE, args: ['--test', 'src/**/*.test.ts']}]},
    // Both projects in one gate, deliberately sequential: two `tsc` processes each
    // re-reading all of src/ are the worst-contending pair in the set, and pitting them
    // against each other cost more than running them back to back (2.1 s + 2.4 s serial
    // vs 4.7 s + 2.6 s when concurrent).
    {
      name: 'typecheck',
      steps: [
        {command: NODE, args: [TSC, '--noEmit', '-p', 'tsconfig.json']},
        {command: NODE, args: [TSC, '--noEmit', '-p', 'tsconfig.test.json']},
      ],
    },
  );
  if (mode === 'full') {
    gates.push(
      {
        name: 'docs:check',
        steps: [
          {command: NODE, args: ['scripts/gen-docs.ts']},
          // --intent-to-add so a brand-new generated page shows up in the diff instead
          // of hiding as untracked; the diff is the assertion that docs match source.
          {command: 'git', args: ['add', '--intent-to-add', '--', 'docs/api']},
          {command: 'git', args: ['diff', '--exit-code', '--', 'docs/api']},
        ],
      },
      {
        // Two cheap invariant checks that share a gate because neither is worth a process slot of
        // its own: both finish in well under a second.
        name: 'invariants',
        steps: [
          {command: NODE, args: ['scripts/check-constitution.ts']},
          {command: NODE, args: ['scripts/check-layering.ts']},
          {command: NODE, args: ['scripts/check-entries.ts']},
        ],
      },
    );
  }
  return gates;
}

/**
 * Run gates concurrently, but only `jobs` at a time. An unbounded fan-out is measurably
 * *slower*: the gates are not competing for cores (14 of them here) but for filesystem
 * throughput — `node --test` already spawns a worker per core, Biome is parallel across
 * all of them, and every gate reads the same few hundred files. Running all of them at
 * once inflated each gate ~2× (29 s of serial work became 57 s of it, for a wall of
 * 19.6 s); two at a time reached the same wall while leaving the machine usable, and
 * beat the wider pools outright in --quick (7.6 s vs 9.9 s). Hence a pool, not a fan-out.
 */
async function runPool(
  gates: Gate[],
  jobs: number,
  report: (result: Result) => void,
): Promise<Result[]> {
  const results: Result[] = [];
  let next = 0;
  const worker = async (): Promise<void> => {
    for (let index = next++; index < gates.length; index = next++) {
      const gate = gates[index];
      if (gate === undefined) return;
      const result = await runGate(gate);
      report(result);
      results.push(result);
    }
  };
  await Promise.all(Array.from({length: Math.min(jobs, gates.length)}, worker));
  return results;
}

async function runGate(gate: Gate): Promise<Result> {
  const started = performance.now();
  for (const step of gate.steps) {
    const {exit, output} = await run(step);
    if (exit !== 0) return {gate, ok: false, ms: performance.now() - started, exit, output};
  }
  return {gate, ok: true, ms: performance.now() - started, exit: 0, output: ''};
}

const seconds = (ms: number) => `${(ms / 1000).toFixed(1)}s`;

/**
 * CI used to enumerate the gates as one workflow step each, purely so a failure named itself in the
 * Checks UI instead of hiding inside a wall of log. That bought one annotation per gate at the price
 * of a second definition of the gate set that nothing kept in step. Workflow commands buy the same
 * annotation from the single definition: an `::error` titled with the gate, and its diagnostics in a
 * collapsed `::group::`. Off outside Actions, where the markers would be noise.
 *
 * @see https://docs.github.com/actions/reference/workflow-commands-for-github-actions
 */
const IN_ACTIONS = process.env.GITHUB_ACTIONS === 'true';

function reportFailure(result: Result): void {
  // Annotation messages are single-line; the escape for a newline is %0A, so the detail goes in the
  // group and the annotation stays a title plus one clause.
  if (IN_ACTIONS) {
    console.error(
      `::error title=verify: ${result.gate.name}::gate failed (exit ${result.exit ?? 'spawn error'})`,
    );
    console.error(`::group::${result.gate.name}`);
  }
  console.error(`\n─── ${result.gate.name} ───\n${result.output.trimEnd()}`);
  if (IN_ACTIONS) console.error('::endgroup::');
}

/**
 * A fingerprint of every byte the gates read that could change their verdict: the commit
 * the tree stands on, the whole working-tree diff against it, and the content of every
 * untracked file git would show. `--binary` so an edited .xlsx fixture contributes its
 * actual bytes rather than a "Binary files differ" placeholder: the key is then complete by
 * construction, and does not rest on reasoning about which diff shapes carry a blob hash.
 *
 * This is what makes a cache hit a *proof* rather than a skip: same key means the gates
 * would be handed the same input, so their answer is already known. Returns undefined when
 * git cannot answer, which disables caching rather than guessing.
 */
async function treeKey(mode: Mode): Promise<string | undefined> {
  const [head, diff, untracked] = await Promise.all([
    run({command: 'git', args: ['rev-parse', 'HEAD']}),
    run({command: 'git', args: ['diff', 'HEAD', '--binary']}),
    run({command: 'git', args: ['ls-files', '--others', '--exclude-standard']}),
  ]);
  if (head.exit !== 0 || diff.exit !== 0 || untracked.exit !== 0) return undefined;

  const hash = createHash('sha256').update(mode).update(head.output).update(diff.output);
  const paths = untracked.output
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .sort();
  for (const path of paths) {
    hash.update(path);
    // A file that vanished between listing and reading is a race, not a verdict: fold in
    // nothing and let the next run — which will see a different tree — decide.
    hash.update(await readFile(resolve(ROOT, path)).catch(() => Buffer.alloc(0)));
  }
  return hash.digest('hex');
}

interface Pass {
  key: string;
  at: string;
}

/** Per-mode, so alternating `--quick --cached` and `--full --cached` don't evict each other. */
async function readPass(mode: Mode): Promise<Pass | undefined> {
  try {
    const parsed: unknown = JSON.parse(await readFile(STAMP, 'utf8'));
    if (typeof parsed !== 'object' || parsed === null) return undefined;
    const pass: unknown = (parsed as Record<string, unknown>)[mode];
    if (typeof pass !== 'object' || pass === null) return undefined;
    const {key, at} = pass as Record<string, unknown>;
    return typeof key === 'string' && typeof at === 'string' ? {key, at} : undefined;
  } catch {
    // Absent or unreadable is a cache miss, never an error: the stamp is derived state.
    return undefined;
  }
}

async function recordPass(mode: Mode, key: string): Promise<void> {
  let stamp: Record<string, Pass> = {};
  try {
    const parsed: unknown = JSON.parse(await readFile(STAMP, 'utf8'));
    if (typeof parsed === 'object' && parsed !== null) stamp = parsed as Record<string, Pass>;
  } catch {
    /* rewritten from scratch */
  }
  stamp[mode] = {key, at: new Date().toISOString()};
  await mkdir(dirname(STAMP), {recursive: true});
  await writeFile(STAMP, `${JSON.stringify(stamp, null, 2)}\n`);
}

async function main() {
  const {mode, list, jobs, cached} = parseArgs(process.argv.slice(2));

  const key = cached ? await treeKey(mode) : undefined;
  if (key !== undefined) {
    const pass = await readPass(mode);
    if (pass?.key === key) {
      console.log(`verify --${mode}: this tree already passed at ${pass.at} — nothing changed`);
      return;
    }
  }

  const gates = await gateSet(mode);

  if (list) {
    for (const gate of gates) console.log(gate.name);
    return;
  }

  const width = Math.max(...gates.map((gate) => gate.name.length));
  console.log(`verify --${mode} — ${gates.length} gates, ${jobs} at a time\n`);

  // Reported as each finishes rather than in declaration order: the fast gates give
  // immediate feedback, and a gate that hangs is identifiable by its absence.
  const started = performance.now();
  const results = await runPool(gates, jobs, (result) => {
    const mark = result.gate.vacuous !== undefined ? '∅' : result.ok ? '✓' : '✗';
    const note =
      result.gate.vacuous !== undefined
        ? `  (${result.gate.vacuous})`
        : result.ok
          ? ''
          : `  exit ${result.exit ?? 'error'}`;
    console.log(
      `  ${mark} ${result.gate.name.padEnd(width)}  ${seconds(result.ms).padStart(5)}${note}`,
    );
  });
  const wall = performance.now() - started;
  const serial = results.reduce((total, result) => total + result.ms, 0);

  const failed = results.filter((result) => !result.ok);
  for (const result of failed) reportFailure(result);

  if (failed.length === 0) {
    if (key !== undefined) await recordPass(mode, key);
    console.log(
      `\nverify: ${gates.length} gates green in ${seconds(wall)} (${seconds(serial)} serial)`,
    );
  } else {
    console.error(
      `\nverify: FAILED — ${failed.map((result) => result.gate.name).join(', ')} ` +
        `(${results.length - failed.length} green) in ${seconds(wall)}`,
    );
    process.exitCode = 1;
  }
}

main().catch((err: unknown) => {
  if (err instanceof UsageError) console.error(`verify: ${err.message}`);
  else {
    console.error(`verify failed: ${err instanceof Error ? (err.stack ?? err.message) : err}`);
  }
  process.exitCode = 1;
});
