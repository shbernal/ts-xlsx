#!/usr/bin/env node
// Regression corpus runner.
//
// Loads every case under cases/, runs each behavior, and reports it green or failed.
//
// It used to be a comparison rather than a verdict. Each behavior recorded a `baseline` — the outcome
// expected of the implementation *at that time* — and the runner crossed it with the actual result to
// distinguish four states: green, a tracked known-open bug, a regression, and a known-open that had
// started passing. That machinery existed because there were two implementations and the corpus
// measured a half-built one against the library it was replacing.
//
// One implementation remained, every one of the 832 behaviors recorded `baseline: 'pass'`, and three of
// the four states became unreachable — the runner carried a comparison whose second operand was a
// constant. So the baseline is gone and this reports what it actually knows: a behavior passed, or it
// failed and the build is red.
//
// Usage:
//   node test/corpus/run.ts [--target src|dist]
//                           [--case <glob>]… [--verbose | --quiet] [--json]
//
//   --case <glob>   run only cases whose `id` or `cluster` matches the glob (`*`/`?`);
//                   repeatable, unioned. A pattern matching nothing is an error, so a
//                   typo can never masquerade as a green run.
//   --verbose       list every behavior, not just the ones that need attention.
//                   Implied by --case, since filtering to a case means wanting to see it.
//   --quiet         force the terse listing even under --case.
//   --json          emit one machine-readable report object on stdout and nothing else.
//
// The summary line is unconditional: it reaches stdout in every mode, filtered or not
// (as the `summary` field under --json), so no invocation can hide the tally.

import assert from 'node:assert/strict';
import {readdir} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import type {Behavior, Case, CorpusApi} from './case.ts';

const HERE = dirname(fileURLToPath(import.meta.url));

/** What a behavior did this run. There is no third state: a corpus behavior runs, or the build is red. */
type Status = 'pass' | 'fail';

interface Outcome {
  status: Status;
  detail?: string;
}

/** One behavior's verdict, as reported. */
interface BehaviorResult {
  name: string;
  status: Status;
  detail?: string;
}

interface Args {
  target: string | undefined;
  /** `--case` globs, unioned; empty means every case. */
  patterns: string[];
  /** Explicit `--verbose`/`--quiet`; unset lets `--case` decide. */
  verbose: boolean | undefined;
  json: boolean;
}

/** A bad invocation, not a corpus failure: reported as one legible line, no stack. */
class UsageError extends Error {}

function parseArgs(argv: string[]): Args {
  const args: Args = {
    target: undefined,
    patterns: [],
    verbose: undefined,
    json: false,
  };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--target') args.target = argv[++i] ?? '';
    else if (argv[i] === '--case') args.patterns.push(argv[++i] ?? '');
    else if (argv[i] === '--verbose') args.verbose = true;
    else if (argv[i] === '--quiet') args.verbose = false;
    else if (argv[i] === '--json') args.json = true;
    else throw new UsageError(`Unrecognized argument: ${argv[i]}`);
  }
  return args;
}

/** Anchored glob over `*` and `?`; every other character is literal. */
function globToRegExp(pattern: string): RegExp {
  const body = pattern.replace(/[.*+?^${}()|[\]\\]/g, (ch) =>
    ch === '*' ? '.*' : ch === '?' ? '.' : `\\${ch}`,
  );
  return new RegExp(`^${body}$`);
}

async function loadCases(): Promise<Case[]> {
  const dir = resolve(HERE, 'cases');
  const files = (await readdir(dir)).filter((f) => f.endsWith('.case.ts')).sort();
  const cases: Case[] = [];
  for (const file of files) {
    const mod = await import(pathToFileURL(resolve(dir, file)).href);
    cases.push(mod.default as Case);
  }
  return cases;
}

/**
 * Narrow to the cases any pattern matches, on `id` or `cluster`. Importing all 265
 * cases costs ~0.2 s against a ~10 s full run, so filtering after load buys nearly
 * the whole saving while matching the case's *durable* identity rather than its
 * filename.
 */
function selectCases(cases: Case[], patterns: string[]): Case[] {
  if (patterns.length === 0) return cases;
  const matchers = patterns.map(globToRegExp);
  const selected = cases.filter((c) => matchers.some((re) => re.test(c.id) || re.test(c.cluster)));
  if (selected.length === 0) {
    // A silent empty run would report green. Fail loudly, and spend a line on the
    // near-misses so a typo costs one invocation instead of two.
    const needle = patterns.join(' ').replace(/[*?]/g, '');
    const near = cases
      .filter((c) => c.id.includes(needle) || c.cluster.includes(needle))
      .map((c) => c.id)
      .slice(0, 8);
    const hint = near.length > 0 ? `\n  did you mean: ${near.join(', ')}` : '';
    throw new UsageError(`no case id or cluster matches ${patterns.join(', ')}${hint}`);
  }
  return selected;
}

async function runBehavior(behavior: Behavior, api: CorpusApi): Promise<Outcome> {
  try {
    await behavior.expect(api, assert);
    return {status: 'pass'};
  } catch (err) {
    return {status: 'fail', detail: err instanceof Error ? err.message : String(err)};
  }
}

const MARK: Record<Status, string> = {pass: '✓', fail: '✗'};

async function main() {
  const args = parseArgs(process.argv.slice(2));
  // Filtering to specific cases is a "show me this one" act, so it lists every
  // behavior by default; an unfiltered run reports only what needs attention,
  // because 756 ✓ lines are noise to every reader we have.
  const verbose = args.verbose ?? args.patterns.length > 0;
  // Under --json stdout carries exactly one object, so nothing else may print there.
  const say = (line: string) => {
    if (!args.json) console.log(line);
  };

  // The adapter picks its target from CORPUS_TARGET at module load, so seed the env before importing
  // it — which is why this import stays dynamic even though there is only one adapter to import. The
  // flag exists because `VAR=value cmd` is POSIX shell syntax that cmd.exe cannot parse, so a package
  // script using it is unrunnable on Windows. Env var and flag stay equivalent; the flag survives the
  // shell.
  if (args.target !== undefined) process.env.CORPUS_TARGET = args.target;
  const adapterPath = resolve(HERE, 'adapters', 'ts-xlsx.ts');
  const adapterMod = (await import(pathToFileURL(adapterPath).href)) as {default: CorpusApi};
  const api = adapterMod.default;
  const cases = selectCases(await loadCases(), args.patterns);

  const tally: Record<Status, number> = {pass: 0, fail: 0};
  const behaviorCount = cases.reduce((n, c) => n + c.behavior.length, 0);
  const filterNote = args.patterns.length > 0 ? ` matching ${args.patterns.join(', ')}` : '';
  say(
    `corpus: ${cases.length} case(s)${filterNote}, ${behaviorCount} behavior(s) vs adapter "${api.name}"\n`,
  );

  const report: {id: string; cluster: string; behaviors: BehaviorResult[]}[] = [];
  for (const testCase of cases) {
    // `provenance` is an optional, disposable trace — a case is identified by its
    // durable `id`/`cluster`, never by an upstream number. Show a ref only if present.
    const ref = testCase.provenance?.ref;
    const heading = `  ${testCase.id}${ref ? `  [#${ref}]` : ''}  ${testCase.cluster}`;
    let headingShown = false;
    const behaviors: BehaviorResult[] = [];

    for (const behavior of testCase.behavior) {
      const {status, detail} = await runBehavior(behavior, api);
      tally[status]++;
      behaviors.push({
        name: behavior.name,
        status,
        ...(detail === undefined ? {} : {detail}),
      });

      const note = status === 'fail' ? `FAILED: ${detail}` : undefined;
      if (!verbose && note === undefined) continue;
      if (!headingShown) {
        say(heading);
        headingShown = true;
      }
      say(`    ${MARK[status]} ${behavior.name}`);
      if (note !== undefined) say(`        ${note}`);
    }

    if (headingShown) say('');
    report.push({id: testCase.id, cluster: testCase.cluster, behaviors});
  }

  const summary = `${tally.pass} green, ${tally.fail} failure(s)`;
  const failed = tally.fail > 0;

  if (args.json) {
    console.log(
      JSON.stringify({
        adapter: api.name,
        target: process.env.CORPUS_TARGET ?? 'src',
        ...(args.patterns.length > 0 ? {patterns: args.patterns} : {}),
        cases: report,
        tally,
        summary,
        ok: !failed,
      }),
    );
  } else {
    say(`summary: ${summary}`);
  }

  if (failed) {
    console.error(`\nFAIL: ${tally.fail} behavior(s) the corpus had locked in no longer hold.`);
    process.exitCode = 1;
  }
}

main().catch((err: unknown) => {
  if (err instanceof UsageError) console.error(`corpus: ${err.message}`);
  else {
    console.error(
      `corpus runner failed: ${err instanceof Error ? (err.stack ?? err.message) : err}`,
    );
  }
  process.exitCode = 1;
});
