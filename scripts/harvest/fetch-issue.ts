#!/usr/bin/env node
// Harvest a single upstream issue or PR into the local backlog queue.
//
// Usage:  node scripts/harvest/fetch-issue.ts <number> [--repo owner/name]
//
// This is the *atom* of the harvest: it turns one GitHub thread into a durable,
// queryable JSON record under docs/knowledge/backlog/issues/, plus any
// spreadsheet-shaped attachments. `harvest-all.ts` fans this out across the
// whole backlog; the shared logic lives in `lib.ts`.

import {DEFAULT_REPO, errorMessage, harvestOne, isValidRepo} from './lib.ts';

interface FetchArgs {
  number: number;
  repo: string;
}

function parseArgs(argv: string[]): FetchArgs {
  let number: number | undefined;
  let repo = DEFAULT_REPO;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i] ?? '';
    if (arg === '--repo') {
      repo = argv[++i] ?? '';
    } else if (/^\d+$/.test(arg)) {
      number = Number(arg);
    } else {
      throw new Error(`Unrecognized argument: ${arg}`);
    }
  }
  if (number === undefined || !Number.isInteger(number) || number <= 0) {
    throw new Error('Expected a positive issue/PR number, e.g. `fetch-issue.ts 140`');
  }
  if (!isValidRepo(repo)) {
    throw new Error(`Invalid --repo value: ${repo}`);
  }
  return {number, repo};
}

async function main() {
  const {number, repo} = parseArgs(process.argv.slice(2));
  const result = await harvestOne({number, repo});

  console.log(`Harvested ${repo}#${number} (${result.type}) -> ${result.recordPath}`);
  console.log(
    `  ${result.commentCount} comment(s), ${result.attachmentCount} link(s), ${result.fixtureCount} fixture candidate(s)`,
  );
  for (const d of result.downloaded ?? []) {
    console.log(d.error ? `  ! ${d.name}: ${d.error}` : `  + ${d.name} (${d.bytes} bytes)`);
  }
}

main().catch((err: unknown) => {
  console.error(`harvest failed: ${errorMessage(err)}`);
  process.exitCode = 1;
});
