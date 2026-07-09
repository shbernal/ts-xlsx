#!/usr/bin/env node
// Harvest a single upstream issue or PR into the local backlog queue.
//
// Usage:  node scripts/harvest/fetch-issue.mjs <number> [--repo owner/name]
//
// This is the *atom* of the harvest: it turns one GitHub thread into a durable,
// queryable JSON record under docs/knowledge/backlog/issues/, plus any
// spreadsheet-shaped attachments. `harvest-all.mjs` fans this out across the
// whole backlog; the shared logic lives in `lib.mjs`. See STRATEGY.md Phase 1
// and docs/knowledge/backlog/README.md.

import {DEFAULT_REPO, harvestOne, isValidRepo} from './lib.mjs';

function parseArgs(argv) {
  const args = {number: undefined, repo: DEFAULT_REPO};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--repo') {
      args.repo = argv[++i];
    } else if (/^\d+$/.test(arg)) {
      args.number = Number(arg);
    } else {
      throw new Error(`Unrecognized argument: ${arg}`);
    }
  }
  if (!Number.isInteger(args.number) || args.number <= 0) {
    throw new Error('Expected a positive issue/PR number, e.g. `fetch-issue.mjs 140`');
  }
  if (!isValidRepo(args.repo)) {
    throw new Error(`Invalid --repo value: ${args.repo}`);
  }
  return args;
}

async function main() {
  const {number, repo} = parseArgs(process.argv.slice(2));
  const result = await harvestOne({number, repo});

  console.log(`Harvested ${repo}#${number} (${result.type}) -> ${result.recordPath}`);
  console.log(
    `  ${result.commentCount} comment(s), ${result.attachmentCount} link(s), ${result.fixtureCount} fixture candidate(s)`
  );
  for (const d of result.downloaded) {
    console.log(d.error ? `  ! ${d.name}: ${d.error}` : `  + ${d.name} (${d.bytes} bytes)`);
  }
}

main().catch(err => {
  console.error(`harvest failed: ${err.message ?? err}`);
  process.exitCode = 1;
});
