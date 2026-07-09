#!/usr/bin/env node
// Follow the harvest: how much is filled, how much remains to drain, what's left.
//
// Usage:  node scripts/harvest/status.mjs [--clusters]
//
// The harvest has two one-way stages against a fixed universe (the manifest):
//   1. FILL  — pull every thread into issues/ (harvest-all.mjs). Runs once.
//   2. DRAIN — distill each thread into durable product (a corpus case and/or
//              spec note) and delete its raw record. The queue counts down to 0.
//
// There is no permanent per-item ledger: a record present in issues/ is work still
// to do; a record absent (once harvestComplete) is work already banked, and its
// commit message is the durable account of what was preserved. This report is the
// live view of that drain — the denominator is the manifest, the numerator is what
// remains on disk.

import {readFile, readdir} from 'node:fs/promises';
import {resolve} from 'node:path';
import {BACKLOG_DIR, ISSUES_DIR, MANIFEST_PATH, REPO_ROOT, fileExists} from './lib.mjs';

const CORPUS_CASES_DIR = resolve(REPO_ROOT, 'test', 'corpus', 'cases');
const SPECS_DIR = resolve(BACKLOG_DIR, '..', 'specs');

function parseArgs(argv) {
  const args = {clusters: false};
  for (const a of argv) {
    if (a === '--clusters') args.clusters = true;
    else throw new Error(`Unrecognized argument: ${a}`);
  }
  return args;
}

async function countFiles(dir, suffix) {
  if (!(await fileExists(dir))) return 0;
  return (await readdir(dir)).filter(f => f.endsWith(suffix)).length;
}

async function presentRecordNumbers() {
  if (!(await fileExists(ISSUES_DIR))) return new Set();
  const files = (await readdir(ISSUES_DIR)).filter(f => /^\d+\.json$/.test(f));
  return new Set(files.map(f => Number(f.replace('.json', ''))));
}

function bar(fraction, width = 32) {
  const filled = Math.round(Math.max(0, Math.min(1, fraction)) * width);
  return `[${'█'.repeat(filled)}${'░'.repeat(width - filled)}]`;
}

function pct(n, d) {
  return d === 0 ? '0%' : `${Math.round((n / d) * 100)}%`;
}

async function main() {
  const {clusters} = parseArgs(process.argv.slice(2));
  const present = await presentRecordNumbers();
  const corpusCases = await countFiles(CORPUS_CASES_DIR, '.case.mjs');
  const specNotes = await countFiles(SPECS_DIR, '.md');

  console.log('\n  ts-xlsx harvest status\n  ──────────────────────');

  if (!(await fileExists(MANIFEST_PATH))) {
    console.log(`  No manifest yet. ${present.size} record(s) harvested ad hoc.`);
    console.log('  Run `node scripts/harvest/list-backlog.mjs` to snapshot the universe, then `harvest-all.mjs`.\n');
    return;
  }

  const manifest = JSON.parse(await readFile(MANIFEST_PATH, 'utf8'));
  const items = manifest.items ?? [];
  const total = manifest.total ?? items.length;
  const inQueue = items.filter(i => present.has(i.number)).length;

  console.log(`  universe: ${total} item(s)  (${manifest.issues ?? '?'} issue(s), ${manifest.pullRequests ?? '?'} PR(s))`);
  console.log(`  captured: ${manifest.generatedAt ?? 'unknown'}`);

  if (!manifest.harvestComplete) {
    const toFetch = total - inQueue;
    console.log(`\n  stage: FILL (queue not yet complete)`);
    console.log(`  ${bar(inQueue / total)}  ${inQueue}/${total} fetched (${pct(inQueue, total)})`);
    console.log(`  ${toFetch} still to fetch — run \`node scripts/harvest/harvest-all.mjs\`.`);
  } else {
    const drained = total - inQueue;
    console.log(`\n  stage: DRAIN`);
    console.log(`  ${bar(drained / total)}  ${drained}/${total} drained (${pct(drained, total)})`);
    console.log(`  remaining in queue: ${inQueue}`);
  }

  console.log(`\n  durable output so far:  ${corpusCases} corpus case(s), ${specNotes} spec note(s)`);

  if (clusters) {
    const remaining = items.filter(i => present.has(i.number));
    const byLabel = new Map();
    for (const it of remaining) {
      const labels = it.labels?.length ? it.labels : ['(unlabeled)'];
      for (const l of labels) byLabel.set(l, (byLabel.get(l) ?? 0) + 1);
    }
    const top = [...byLabel.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15);
    if (top.length) {
      console.log(`\n  remaining queue by label (top ${top.length}):`);
      const width = Math.max(...top.map(([l]) => l.length));
      for (const [label, n] of top) console.log(`    ${label.padEnd(width)}  ${n}`);
    }
  } else {
    console.log(`  (pass --clusters to see the remaining queue broken down by label)`);
  }
  console.log('');
}

main().catch(err => {
  console.error(`status failed: ${err.message ?? err}`);
  process.exitCode = 1;
});
