#!/usr/bin/env node
// Coverage measured over *both* suites, because either one alone reports a number that is wrong.
//
// The library is tested by two deliberately separate suites (docs/architecture.md): the white-box
// unit tests co-located as `src/**/*.test.ts`, and the implementation-blind black-box corpus under
// `test/corpus/`, which reaches the library only through an adapter. They are complementary by
// design — that is the whole point of the topology — so neither one's coverage is the library's
// coverage.
//
// The old `test:coverage` measured only the unit suite. That did not merely under-report; it
// reported specific, confident, wrong numbers about modules the corpus exercises in full:
//
//     src/core/table-style.ts    81.25 % lines,  0.00 % functions   →  union: 97.22 / 100.00
//     src/io/xlsb/formula.ts     77.69 % lines, 66.67 % functions   →  union: 97.31 / 100.00
//
// A zero that means "the other suite covers this" is indistinguishable from a zero that means
// "nothing tests this", so the report cost more than it bought: a reader chasing the worst rows in
// the table spends their time on modules that were never uncovered, while a genuinely untested
// function sits mid-table looking fine. The one real gap this repo had — `iconSetXml`, a documented
// rule type nothing ever read back — was invisible in that table for exactly that reason.
//
// So: run each suite with `NODE_V8_COVERAGE` pointed at one shared directory, and report the union.
// V8 writes one JSON file per process, keyed by pid and timestamp, and node's own coverage reader
// merges every `coverage-*.json` it finds in the directory. The suites need no knowledge of each
// other and stay separate processes, which keeps the corpus implementation-blind.
//
// Usage:
//   node scripts/coverage.ts [--suite <name>]… [--reuse]
//                            [--lines <pct>] [--branches <pct>] [--functions <pct>]
//
//   --suite <name>    measure only this suite (`unit` or `corpus`); repeatable, unioned.
//                     Defaults to every suite — which is the only total that is true.
//                     Useful to see what one suite contributes, never to judge the library.
//   --reuse           report from the raw coverage already in .tmp/coverage without re-running
//                     anything. Re-renders in milliseconds; reports stale numbers if the tree moved.
//   --lines           floors, as percentages. A total below its floor exits non-zero.
//   --branches
//   --functions
//
// ## Why this borrows node's own coverage implementation
//
// The numbers have to be *comparable* to `node --test --experimental-test-coverage`, or this tool
// replaces one misleading report with another: a reader seeing a different total could not tell
// whether the delta was the corpus or this script's arithmetic. Line coverage is not a quantity you
// can eyeball — it is byte-offset ranges mapped onto lines, with `/* node:coverage ignore */`
// handling, block-coverage branch counting, and a specific rule for which lines count at all.
// Reimplementing that produces numbers that are *close*, and close is the failure mode being fixed.
//
// So we construct node's own `TestCoverage` against our directory and call its `summary()`. This is
// an internal module and needs `--expose-internals`, which is why the report phase re-executes this
// script with that flag. That coupling is deliberate and bounded: it fails *loudly* — the require
// throws `Cannot find module` and this script stops with the message below — rather than quietly
// producing numbers that no longer mean what they say. Given the failure being fixed is precisely
// "confident wrong numbers", a loud break on a node upgrade is the cheaper risk. See ADR 0035.

import {spawn} from 'node:child_process';
import {mkdir, readFile, rm, writeFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import {dirname, join, relative, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const RAW = join(ROOT, '.tmp', 'coverage');
const NODE = process.execPath;
const SELF = fileURLToPath(import.meta.url);

/** Set on the re-executed child so a node that no longer exposes the internals fails instead of forking forever. */
const RESPAWNED = 'TS_XLSX_COVERAGE_RESPAWNED';

/**
 * The library is what gets measured. Test files, the corpus harness, and `scripts/` are the
 * instrument, and an instrument that reports on itself flatters itself.
 */
const INCLUDE = ['src/**/*.ts'];
const EXCLUDE = ['src/**/*.test.ts'];

/**
 * Floors, not targets (CLAUDE.md §2). Set just under the measured union so a real regression trips
 * them and ordinary churn does not; raise them when the true number rises, never lower them to fit
 * a change. Totals only, deliberately: a per-file floor would have to be either so low it proves
 * nothing or so exact it fails on every honest refactor, and the per-file table below is already
 * the place to see a single module slipping.
 */
const FLOORS = {lines: 98, branches: 92, functions: 98} as const;

interface Suite {
  name: string;
  args: string[];
}

const SUITES: readonly Suite[] = [
  {name: 'unit', args: ['--test', 'src/**/*.test.ts']},
  {name: 'corpus', args: ['test/corpus/run.ts']},
];

/** A bad invocation, not a failing check: one legible line, no stack. */
class UsageError extends Error {}

/** The internals are gone or renamed — the one failure this tool must never paper over. */
class InternalsUnavailableError extends Error {}

interface Args {
  suites: readonly Suite[];
  reuse: boolean;
  floors: {lines: number; branches: number; functions: number};
}

function parsePercent(raw: string | undefined, flag: string): number {
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0 || value > 100) {
    throw new UsageError(`${flag} needs a percentage between 0 and 100, got ${raw}`);
  }
  return value;
}

function parseArgs(argv: readonly string[]): Args {
  const names: string[] = [];
  const args: Args = {suites: SUITES, reuse: false, floors: {...FLOORS}};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--reuse') args.reuse = true;
    else if (arg === '--suite') {
      const name = argv[++i];
      const suite = SUITES.find((candidate) => candidate.name === name);
      if (suite === undefined) {
        throw new UsageError(
          `--suite needs one of ${SUITES.map((s) => s.name).join(', ')}, got ${name}`,
        );
      }
      names.push(suite.name);
    } else if (arg === '--lines') args.floors.lines = parsePercent(argv[++i], '--lines');
    else if (arg === '--branches') args.floors.branches = parsePercent(argv[++i], '--branches');
    else if (arg === '--functions') args.floors.functions = parsePercent(argv[++i], '--functions');
    else {
      throw new UsageError(
        `unrecognized argument: ${arg} ` +
          `(expected --suite, --reuse, --lines, --branches, --functions)`,
      );
    }
  }
  if (names.length > 0) args.suites = SUITES.filter((suite) => names.includes(suite.name));
  return args;
}

/**
 * Rebuild the command line for the re-executed child, which must not re-run the suites. `--suite`
 * is deliberately not forwarded: the child learns what it is reporting on from the manifest in
 * {@link RAW}, so there is no second, flag-shaped account of the scope that could disagree with it.
 */
function toArgv(args: Args): string[] {
  return [
    '--reuse',
    '--lines',
    String(args.floors.lines),
    '--branches',
    String(args.floors.branches),
    '--functions',
    String(args.floors.functions),
  ];
}

function run(command: string, argv: string[], env?: NodeJS.ProcessEnv): Promise<number | null> {
  return new Promise((settle) => {
    const child = spawn(command, argv, {
      cwd: ROOT,
      // Inherited, not captured: a suite that fails must show why, and a coverage run long enough
      // to wonder about should show progress while it happens.
      stdio: 'inherit',
      env: {...process.env, ...env},
    });
    child.on('error', () => settle(null));
    child.on('close', (code) => settle(code));
  });
}

/**
 * Which suites produced the raw coverage now in {@link RAW}. Written beside it rather than inferred
 * from the flags, because under `--reuse` the flags describe an invocation that did not happen: a
 * `--suite unit --reuse` after a full run would otherwise label the union as unit-only, which is the
 * same species of confident-wrong label this whole tool exists to remove. The name cannot collide
 * with the `coverage-*.json` files node reads out of the directory.
 */
const MANIFEST = join(RAW, 'suites.json');

async function readManifest(): Promise<string[] | undefined> {
  try {
    const parsed: unknown = JSON.parse(await readFile(MANIFEST, 'utf8'));
    return Array.isArray(parsed) && parsed.every((name) => typeof name === 'string')
      ? parsed
      : undefined;
  } catch {
    // Absent or unreadable means we cannot name the scope, never that the scope is everything.
    return undefined;
  }
}

/**
 * Run every requested suite into one directory, cleared first. Stale files from an earlier run
 * would merge in silently and inflate the result — the same class of quiet wrongness this tool
 * exists to remove.
 */
async function runSuites(suites: readonly Suite[]): Promise<void> {
  await rm(RAW, {recursive: true, force: true});
  await mkdir(RAW, {recursive: true});
  for (const suite of suites) {
    const started = performance.now();
    const code = await run(NODE, suite.args, {NODE_V8_COVERAGE: RAW});
    const elapsed = ((performance.now() - started) / 1000).toFixed(1);
    if (code !== 0) {
      throw new Error(
        `suite '${suite.name}' failed (exit ${code ?? 'spawn error'}) — ` +
          `coverage of a red tree measures nothing`,
      );
    }
    console.log(`coverage: ${suite.name} ran green in ${elapsed}s`);
  }
  await writeFile(MANIFEST, `${JSON.stringify(suites.map((suite) => suite.name))}\n`);
}

// The shape of the two internal entry points we lean on, declared rather than inferred: an internal
// module carries no types, and `any` here would let a silent shape change through as `undefined`.
interface CoverageTotals {
  coveredLinePercent: number;
  coveredBranchPercent: number;
  coveredFunctionPercent: number;
}

interface CoverageSummary {
  files: {path: string}[];
  totals: CoverageTotals;
}

interface CoverageOptions {
  cwd: string;
  lineCoverage: number;
  branchCoverage: number;
  functionCoverage: number;
  sourceMaps: boolean;
  coverageIncludeGlobs: readonly string[];
  coverageExcludeGlobs: readonly string[];
}

interface Internals {
  TestCoverage: new (
    coverageDirectory: string,
    originalCoverageDirectory: string | undefined,
    options: CoverageOptions,
  ) => {summary(): CoverageSummary};
  getCoverageReport: (
    pad: string,
    summary: CoverageSummary,
    symbol: string,
    color: string,
    table: boolean,
  ) => string;
}

/** Undefined when this node was not started with `--expose-internals`; throws when they are gone. */
function loadInternals(): Internals | undefined {
  const require = createRequire(import.meta.url);
  let coverage: unknown;
  let utils: unknown;
  try {
    coverage = require('internal/test_runner/coverage');
    utils = require('internal/test_runner/utils');
  } catch {
    if (process.env[RESPAWNED] === undefined) return undefined;
    throw new InternalsUnavailableError(
      `node ${process.version} does not expose internal/test_runner/coverage under ` +
        `--expose-internals. This script borrows node's own coverage implementation so its ` +
        `numbers match \`node --test --experimental-test-coverage\` exactly (ADR 0035); that ` +
        `module has moved or gone. Fix the import or reimplement summary() — do not fall back ` +
        `to an approximation, which is the failure this tool exists to remove.`,
    );
  }
  const {TestCoverage} = coverage as {TestCoverage?: unknown};
  const {getCoverageReport} = utils as {getCoverageReport?: unknown};
  if (typeof TestCoverage !== 'function' || typeof getCoverageReport !== 'function') {
    throw new InternalsUnavailableError(
      `node ${process.version} exposes internal/test_runner but not the expected ` +
        `TestCoverage / getCoverageReport shape (ADR 0035).`,
    );
  }
  return {TestCoverage, getCoverageReport} as Internals;
}

/**
 * The `src/` modules absent from the report. V8 records only scripts it actually loaded, so a module
 * no suite imports does not appear as 0 % — it does not appear at all, which reads as "no problem
 * here". Today every one of these is a type-only module or a re-export barrel that erases to nothing
 * at runtime (their contracts are gated by `typecheck` and `check-entries.ts` instead), so the list
 * is expected and short. A module with real behavior showing up here is the signal.
 */
function unmeasured(summary: CoverageSummary, sources: readonly string[]): string[] {
  const measured = new Set(
    summary.files.map((file) => relative(ROOT, file.path).split('\\').join('/')),
  );
  return sources.filter((source) => !measured.has(source));
}

async function listSources(): Promise<string[]> {
  return new Promise((settle) => {
    const child = spawn('git', ['ls-files', 'src'], {
      cwd: ROOT,
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    const chunks: Buffer[] = [];
    child.stdout.on('data', (chunk: Buffer) => chunks.push(chunk));
    child.on('error', () => settle([]));
    child.on('close', () => {
      settle(
        Buffer.concat(chunks)
          .toString('utf8')
          .split('\n')
          .map((line) => line.trim())
          .filter((line) => line.endsWith('.ts') && !line.endsWith('.test.ts')),
      );
    });
  });
}

function summarize(internals: Internals): CoverageSummary {
  return new internals.TestCoverage(RAW, undefined, {
    cwd: ROOT,
    // Node's own thresholds only tint the table; the verdict below is ours, so these stay at zero.
    lineCoverage: 0,
    branchCoverage: 0,
    functionCoverage: 0,
    sourceMaps: false,
    coverageIncludeGlobs: INCLUDE,
    coverageExcludeGlobs: EXCLUDE,
  }).summary();
}

/** Names the data actually in {@link RAW}, and says plainly when that data is not the library. */
function describeScope(measured: readonly string[] | undefined): string {
  if (measured === undefined) return 'an unrecorded set of suites — re-run without --reuse';
  if (measured.length === SUITES.length) return 'unit + corpus';
  return `${measured.join(' + ')} only — NOT the library's coverage`;
}

function report(
  summary: CoverageSummary,
  internals: Internals,
  args: Args,
  measured: readonly string[] | undefined,
): void {
  console.log(internals.getCoverageReport('', summary, '', '', true));

  const {totals} = summary;
  const checks = [
    {label: 'lines', actual: totals.coveredLinePercent, floor: args.floors.lines},
    {label: 'branches', actual: totals.coveredBranchPercent, floor: args.floors.branches},
    {label: 'functions', actual: totals.coveredFunctionPercent, floor: args.floors.functions},
  ];
  // A partial run is measuring one instrument, not the library, so its totals have nothing to be
  // held to. Floors apply to the union or to nothing.
  const enforced = measured?.length === SUITES.length;
  const below = enforced ? checks.filter((check) => check.actual < check.floor) : [];

  console.log(`coverage: ${summary.files.length} modules measured over ${describeScope(measured)}`);

  if (below.length > 0) {
    for (const check of below) {
      console.error(
        `coverage: FAILED — ${check.label} ${check.actual.toFixed(2)}% is below the ` +
          `${check.floor}% floor`,
      );
    }
    process.exitCode = 1;
  }
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  if (!args.reuse) await runSuites(args.suites);

  const internals = loadInternals();
  if (internals === undefined) {
    const code = await run(NODE, ['--expose-internals', SELF, ...toArgv(args)], {[RESPAWNED]: '1'});
    process.exitCode = code ?? 1;
    return;
  }

  const summary = summarize(internals);
  report(summary, internals, args, await readManifest());

  const missing = unmeasured(summary, await listSources());
  if (missing.length > 0) {
    console.log(
      `coverage: ${missing.length} modules contributed no runtime code and are absent from the ` +
        `table (expected for type-only modules and re-export barrels):`,
    );
    for (const file of missing) console.log(`  ${file}`);
  }
}

main().catch((err: unknown) => {
  if (err instanceof UsageError) console.error(`coverage: ${err.message}`);
  else if (err instanceof InternalsUnavailableError) console.error(`coverage: ${err.message}`);
  else
    console.error(
      `coverage failed: ${err instanceof Error ? (err.stack ?? err.message) : String(err)}`,
    );
  process.exitCode = 1;
});
