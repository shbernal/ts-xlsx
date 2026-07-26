#!/usr/bin/env node
// Regression corpus runner.
//
// Loads every case under cases/, runs each behavior against an adapter (default:
// the rewrite), and compares the *actual* outcome to the behavior's recorded
// `baseline` — the outcome we expect the implementation to produce.
//
//   baseline  actual   meaning                          fails build?
//   --------  ------   ------------------------------    ------------
//   pass      pass     green, as expected               no
//   fail      fail     known-open bug, tracked          no   (an intentional pending marker)
//   pass      fail     REGRESSION                       YES
//   fail      pass     bug fixed — flip baseline to pass no   (but loudly flagged)
//
// Post-Phase-3 the rewrite is the reference implementation and every behavior in the
// corpus baselines to `pass`. The `baseline: 'fail'` marker survives as the way to
// land a case for a not-yet-built capability without reddening CI — a tracked
// known-open, not a legacy oracle's verdict.
//
// Usage:
//   node test/corpus/run.ts [--adapter rewrite] [--target src|dist]
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

/** What a behavior actually did this run. */
type Actual = 'pass' | 'fail' | 'skip';
/** How that outcome reads against its baseline. */
type Status = 'ok' | 'bug' | 'regression' | 'fixed' | 'skip';

interface Outcome {
  actual: Actual;
  detail?: string;
}

/** One behavior's verdict, as reported. */
interface BehaviorResult {
  name: string;
  baseline: 'pass' | 'fail';
  actual: Actual;
  status: Status;
  detail?: string;
}

interface Args {
  adapter: string;
  target: string | undefined;
  /** `--case` globs, unioned; empty means every case. */
  patterns: string[];
  /** Explicit `--verbose`/`--quiet`; unset lets `--case` decide. */
  verbose: boolean | undefined;
  json: boolean;
}

/** A bad invocation, not a corpus failure: reported as one legible line, no stack. */
class UsageError extends Error {}

// An adapter that does not yet implement a capability tags its error object so the
// behavior is SKIPPED rather than counted as a failure/regression.
function isNotImplemented(err: unknown): boolean {
  return (
    typeof err === 'object' &&
    err !== null &&
    'notImplemented' in err &&
    Boolean((err as {notImplemented?: unknown}).notImplemented)
  );
}

function parseArgs(argv: string[]): Args {
  const args: Args = {
    adapter: 'rewrite',
    target: undefined,
    patterns: [],
    verbose: undefined,
    json: false,
  };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--adapter') args.adapter = argv[++i] ?? '';
    else if (argv[i] === '--target') args.target = argv[++i] ?? '';
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
    return {actual: 'pass'};
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    if (isNotImplemented(err)) return {actual: 'skip', detail};
    return {actual: 'fail', detail};
  }
}

function classify(baseline: 'pass' | 'fail', actual: Actual): Status {
  if (actual === 'skip') return 'skip';
  if (baseline === 'pass') return actual === 'pass' ? 'ok' : 'regression';
  return actual === 'fail' ? 'bug' : 'fixed';
}

const MARK: Record<Status, string> = {ok: '✓', bug: '○', regression: '✗', fixed: '↑', skip: '∅'};

/** The extra prose a non-green status owes the reader. */
function annotate(status: Status, detail: string | undefined): string | undefined {
  if (status === 'regression') return `REGRESSION: ${detail}`;
  if (status === 'bug') return `known-open (baseline=fail): ${detail}`;
  if (status === 'fixed') return `FIXED — now passes; flip this behavior's baseline to 'pass'`;
  if (status === 'skip') return `not implemented: ${detail}`;
  return undefined;
}

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

  // The adapter picks its target from CORPUS_TARGET at module load, so seed the env
  // before importing it. The flag exists because `VAR=value cmd` is POSIX shell syntax
  // that cmd.exe cannot parse — a package script using it is unrunnable on Windows.
  // Env var and flag stay equivalent; the flag simply survives the shell.
  if (args.target !== undefined) process.env.CORPUS_TARGET = args.target;
  const adapterPath = resolve(HERE, 'adapters', `${args.adapter}.ts`);
  const adapterMod = await import(pathToFileURL(adapterPath).href);
  const api = adapterMod.default;
  const cases = selectCases(await loadCases(), args.patterns);

  const tally: Record<Status, number> = {ok: 0, bug: 0, regression: 0, fixed: 0, skip: 0};
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
      const {actual, detail} = await runBehavior(behavior, api);
      const status = classify(behavior.baseline, actual);
      tally[status]++;
      behaviors.push({
        name: behavior.name,
        baseline: behavior.baseline,
        actual,
        status,
        ...(detail === undefined ? {} : {detail}),
      });

      const note = annotate(status, detail);
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

  const skipNote =
    tally.skip > 0 ? `, ${tally.skip} skipped (capability not implemented by "${api.name}")` : '';
  const summary = `${tally.ok} green, ${tally.bug} known-open, ${tally.fixed} newly-fixed, ${tally.regression} regression(s)${skipNote}`;
  const failed = tally.regression > 0;

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
    console.error('\nFAIL: regression(s) detected — a behavior that passed on legacy now fails.');
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
