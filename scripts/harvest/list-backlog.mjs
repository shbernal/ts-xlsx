#!/usr/bin/env node
// Snapshot the entire open upstream backlog into a manifest — done ONCE.
//
// Usage:  node scripts/harvest/list-backlog.mjs [--repo owner/name]
//
// The harvest is a one-time operation (STRATEGY.md Phase 1): this manifest is the
// proof-of-universe. It records every open issue + PR that existed at fork time so
// the drain has a fixed denominator — `status.mjs` measures progress against it,
// and it lets us prove nothing was silently dropped (CLAUDE.md §"no silent caps")
// even though we delete each raw record as we process it.
//
// We do NOT re-run this to pick up new upstream activity; the universe is frozen
// the moment we capture it. Re-running only refreshes titles/labels/counts for the
// same frozen set and preserves the `harvestComplete` flag.

import {readFile, writeFile, mkdir} from 'node:fs/promises';
import {dirname} from 'node:path';
import {BACKLOG_DIR, DEFAULT_REPO, MANIFEST_PATH, fileExists, isValidRepo, listOpenItems} from './lib.mjs';

function parseArgs(argv) {
  const args = {repo: DEFAULT_REPO};
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--repo') args.repo = argv[++i];
    else throw new Error(`Unrecognized argument: ${argv[i]}`);
  }
  if (!isValidRepo(args.repo)) throw new Error(`Invalid --repo value: ${args.repo}`);
  return args;
}

async function main() {
  const {repo} = parseArgs(process.argv.slice(2));

  // Preserve the drain marker across a refresh — re-listing must not silently
  // "un-complete" a finished harvest.
  let harvestComplete = false;
  if (await fileExists(MANIFEST_PATH)) {
    try {
      harvestComplete = JSON.parse(await readFile(MANIFEST_PATH, 'utf8')).harvestComplete ?? false;
    } catch {
      // A corrupt manifest is rewritten from scratch.
    }
  }

  console.log(`Enumerating open issues + PRs in ${repo} …`);
  const items = await listOpenItems(repo);
  items.sort((a, b) => a.number - b.number);

  const issues = items.filter(i => i.type === 'issue').length;
  const prs = items.length - issues;

  const manifest = {
    schema: 'ts-xlsx/backlog-manifest@1',
    repo,
    generatedAt: new Date().toISOString(),
    total: items.length,
    issues,
    pullRequests: prs,
    harvestComplete,
    items,
  };

  await mkdir(dirname(MANIFEST_PATH), {recursive: true});
  await writeFile(MANIFEST_PATH, `${JSON.stringify(manifest, null, 2)}\n`);

  console.log(`Wrote manifest: ${MANIFEST_PATH}`);
  console.log(`  universe: ${items.length} item(s) — ${issues} issue(s), ${prs} PR(s)`);
  console.log(`Next: node scripts/harvest/harvest-all.mjs   (fill the queue)`);
}

main().catch(err => {
  console.error(`list-backlog failed: ${err.message ?? err}`);
  process.exitCode = 1;
});
