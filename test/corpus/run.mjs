#!/usr/bin/env node
// Regression corpus runner.
//
// Loads every case under cases/, runs each behavior against an adapter (default:
// the current legacy code), and compares the *actual* outcome to the behavior's
// recorded `baseline` — the outcome we know today's legacy code produces.
//
//   baseline  actual   meaning                         fails build?
//   --------  ------   -----------------------------   ------------
//   pass      pass     green, as expected              no
//   fail      fail     known-open bug, tracked         no   (this is the point of the corpus)
//   pass      fail     REGRESSION                      YES
//   fail      pass     bug fixed — update the baseline no   (but loudly flagged)
//
// This is what lets the corpus run against the frozen legacy tree "mostly red where
// bugs are real" (STRATEGY.md Phase 1) without a red build, while still catching any
// real regression the moment it appears. The rewrite's bar is: every baseline flips
// to `pass`.
//
// Usage:  node test/corpus/run.mjs [--adapter current]

import assert from 'node:assert/strict';
import {readdir} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const args = {adapter: 'current'};
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

function runBehavior(behavior, api) {
  try {
    behavior.expect(api, assert);
    return {actual: 'pass'};
  } catch (err) {
    return {actual: 'fail', detail: err.message ?? String(err)};
  }
}

const MARK = {ok: '✓', bug: '○', regression: '✗', fixed: '↑'};

async function main() {
  const {adapter: adapterName} = parseArgs(process.argv.slice(2));
  const adapterMod = await import(pathToFileURL(resolve(HERE, 'adapters', `${adapterName}.mjs`)).href);
  const api = adapterMod.default;
  const cases = await loadCases();

  const tally = {ok: 0, bug: 0, regression: 0, fixed: 0};
  const behaviorCount = cases.reduce((n, c) => n + c.behavior.length, 0);
  console.log(`corpus: ${cases.length} case(s), ${behaviorCount} behavior(s) vs adapter "${api.name}"\n`);

  for (const testCase of cases) {
    console.log(`  ${testCase.id}  [#${testCase.provenance.ref}]  ${testCase.cluster}`);
    for (const behavior of testCase.behavior) {
      const {actual, detail} = runBehavior(behavior, api);
      let status;
      if (behavior.baseline === 'pass' && actual === 'pass') status = 'ok';
      else if (behavior.baseline === 'fail' && actual === 'fail') status = 'bug';
      else if (behavior.baseline === 'pass' && actual === 'fail') status = 'regression';
      else status = 'fixed';
      tally[status]++;
      console.log(`    ${MARK[status]} ${behavior.name}`);
      if (status === 'regression') console.log(`        REGRESSION: ${detail}`);
      if (status === 'bug') console.log(`        known-open (baseline=fail): ${detail}`);
      if (status === 'fixed') console.log(`        FIXED — flip baseline to 'pass' for this behavior`);
    }
    console.log('');
  }

  console.log(
    `summary: ${tally.ok} green, ${tally.bug} known-open, ${tally.fixed} newly-fixed, ${tally.regression} regression(s)`
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
