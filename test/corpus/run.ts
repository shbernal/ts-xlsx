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
// Usage:  node test/corpus/run.ts [--adapter rewrite]

import assert from 'node:assert/strict';
import {access, readdir} from 'node:fs/promises';
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

// Resolve a module that may be mid-migration from `.mjs` to `.ts`, preferring the
// migrated `.ts` when both exist.
async function resolveExisting(...candidates: string[]): Promise<string> {
  for (const path of candidates) {
    try {
      await access(path);
      return path;
    } catch {}
  }
  throw new Error(`none of these exist:\n  ${candidates.join('\n  ')}`);
}

function parseArgs(argv: string[]): {adapter: string} {
  const args = {adapter: 'rewrite'};
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--adapter') args.adapter = argv[++i] ?? '';
    else throw new Error(`Unrecognized argument: ${argv[i]}`);
  }
  return args;
}

async function loadCases(): Promise<Case[]> {
  const dir = resolve(HERE, 'cases');
  // Cases are migrating `.case.mjs` → `.case.ts`; accept either so the rename can
  // land incrementally without a flag day.
  const files = (await readdir(dir))
    .filter((f) => f.endsWith('.case.mjs') || f.endsWith('.case.ts'))
    .sort();
  const cases: Case[] = [];
  for (const file of files) {
    const mod = await import(pathToFileURL(resolve(dir, file)).href);
    cases.push(mod.default as Case);
  }
  return cases;
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

const MARK: Record<Status, string> = {ok: '✓', bug: '○', regression: '✗', fixed: '↑', skip: '∅'};

async function main() {
  const {adapter: adapterName} = parseArgs(process.argv.slice(2));
  const adapterPath = await resolveExisting(
    resolve(HERE, 'adapters', `${adapterName}.ts`),
    resolve(HERE, 'adapters', `${adapterName}.mjs`),
  );
  const adapterMod = await import(pathToFileURL(adapterPath).href);
  const api = adapterMod.default;
  const cases = await loadCases();

  const tally: Record<Status, number> = {ok: 0, bug: 0, regression: 0, fixed: 0, skip: 0};
  const behaviorCount = cases.reduce((n, c) => n + c.behavior.length, 0);
  console.log(
    `corpus: ${cases.length} case(s), ${behaviorCount} behavior(s) vs adapter "${api.name}"\n`,
  );

  for (const testCase of cases) {
    // `provenance` is an optional, disposable trace — a case is identified by its
    // durable `id`/`cluster`, never by an upstream number. Show a ref only if present.
    const ref = testCase.provenance?.ref;
    const tag = ref ? `  [#${ref}]` : '';
    console.log(`  ${testCase.id}${tag}  ${testCase.cluster}`);
    for (const behavior of testCase.behavior) {
      const {actual, detail} = await runBehavior(behavior, api);
      let status: Status;
      if (actual === 'skip') status = 'skip';
      else if (behavior.baseline === 'pass' && actual === 'pass') status = 'ok';
      else if (behavior.baseline === 'fail' && actual === 'fail') status = 'bug';
      else if (behavior.baseline === 'pass' && actual === 'fail') status = 'regression';
      else status = 'fixed';
      tally[status]++;
      console.log(`    ${MARK[status]} ${behavior.name}`);
      if (status === 'regression') console.log(`        REGRESSION: ${detail}`);
      if (status === 'bug') console.log(`        known-open (baseline=fail): ${detail}`);
      if (status === 'fixed') {
        console.log(`        FIXED — now passes; flip this behavior's baseline to 'pass'`);
      }
    }
    console.log('');
  }

  const skipNote =
    tally.skip > 0 ? `, ${tally.skip} skipped (capability not implemented by "${api.name}")` : '';
  console.log(
    `summary: ${tally.ok} green, ${tally.bug} known-open, ${tally.fixed} newly-fixed, ${tally.regression} regression(s)${skipNote}`,
  );
  if (tally.regression > 0) {
    console.error('\nFAIL: regression(s) detected — a behavior that passed on legacy now fails.');
    process.exitCode = 1;
  }
}

main().catch((err: unknown) => {
  console.error(`corpus runner failed: ${err instanceof Error ? (err.stack ?? err.message) : err}`);
  process.exitCode = 1;
});
