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
// Usage:  node test/corpus/run.mjs [--adapter rewrite]

import assert from 'node:assert/strict';
import {readdir} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const args = {adapter: 'rewrite'};
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--adapter') args.adapter = argv[++i];
    else throw new Error(`Unrecognized argument: ${argv[i]}`);
  }
  return args;
}

async function loadCases() {
  const dir = resolve(HERE, 'cases');
  const files = (await readdir(dir)).filter(f => f.endsWith('.case.mjs')).sort();
  const cases = [];
  for (const file of files) {
    const mod = await import(pathToFileURL(resolve(dir, file)).href);
    cases.push(mod.default);
  }
  return cases;
}

async function runBehavior(behavior, api) {
  try {
    await behavior.expect(api, assert);
    return {actual: 'pass'};
  } catch (err) {
    // An adapter that does not yet implement a capability (the in-progress rewrite)
    // tags the error so the behavior is SKIPPED, not counted as a failure/regression.
    if (err?.notImplemented) return {actual: 'skip', detail: err.message ?? String(err)};
    return {actual: 'fail', detail: err.message ?? String(err)};
  }
}

const MARK = {ok: '✓', bug: '○', regression: '✗', fixed: '↑', skip: '∅'};

async function main() {
  const {adapter: adapterName} = parseArgs(process.argv.slice(2));
  const adapterMod = await import(pathToFileURL(resolve(HERE, 'adapters', `${adapterName}.mjs`)).href);
  const api = adapterMod.default;
  const cases = await loadCases();

  const tally = {ok: 0, bug: 0, regression: 0, fixed: 0, skip: 0};
  const behaviorCount = cases.reduce((n, c) => n + c.behavior.length, 0);
  console.log(`corpus: ${cases.length} case(s), ${behaviorCount} behavior(s) vs adapter "${api.name}"\n`);

  for (const testCase of cases) {
    // `provenance` is an optional, disposable trace — a case is identified by its
    // durable `id`/`cluster`, never by an upstream number. Show a ref only if present.
    const ref = testCase.provenance?.ref;
    const tag = ref ? `  [#${ref}]` : '';
    console.log(`  ${testCase.id}${tag}  ${testCase.cluster}`);
    for (const behavior of testCase.behavior) {
      const {actual, detail} = await runBehavior(behavior, api);
      let status;
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

  const skipNote = tally.skip > 0 ? `, ${tally.skip} skipped (capability not implemented by "${api.name}")` : '';
  console.log(
    `summary: ${tally.ok} green, ${tally.bug} known-open, ${tally.fixed} newly-fixed, ${tally.regression} regression(s)${skipNote}`
  );
  if (tally.regression > 0) {
    console.error('\nFAIL: regression(s) detected — a behavior that passed on legacy now fails.');
    process.exitCode = 1;
  }
}

main().catch(err => {
  console.error(`corpus runner failed: ${err.stack ?? err}`);
  process.exitCode = 1;
});
