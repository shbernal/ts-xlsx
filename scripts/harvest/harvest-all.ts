#!/usr/bin/env node
// Fill the backlog queue: fetch every item in the manifest into issues/.
//
// Usage:  node scripts/harvest/harvest-all.ts [--repo owner/name] [--concurrency N]
//
// Reads the manifest written by list-backlog.ts and harvests each thread that
// isn't already on disk. Resumable by design — rerun after a failure or interrupt
// and it skips what it already has and retries the rest. When every item is
// present it flips the manifest's `harvestComplete` flag, which is how status.ts
// knows an absent record means "drained" rather than "never fetched".
//
// This is the one-time bulk fill. After it, agents drain the queue thread by
// thread — see the harvest-triage skill.

import {readFile, writeFile} from 'node:fs/promises';
import type {Manifest, ManifestItem} from './lib.ts';
import {
  DEFAULT_REPO,
  errorMessage,
  fileExists,
  harvestOne,
  isValidRepo,
  MANIFEST_PATH,
  recordPathFor,
} from './lib.ts';

interface HarvestAllArgs {
  repo: string;
  concurrency: number;
}

function parseArgs(argv: string[]): HarvestAllArgs {
  const args: HarvestAllArgs = {repo: DEFAULT_REPO, concurrency: 6};
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--repo') args.repo = argv[++i] ?? '';
    else if (argv[i] === '--concurrency') args.concurrency = Number(argv[++i]);
    else throw new Error(`Unrecognized argument: ${argv[i]}`);
  }
  if (!isValidRepo(args.repo)) throw new Error(`Invalid --repo value: ${args.repo}`);
  if (!Number.isInteger(args.concurrency) || args.concurrency < 1 || args.concurrency > 16) {
    throw new Error('--concurrency must be an integer in [1, 16]');
  }
  return args;
}

// A bounded worker pool: `concurrency` workers pull from a shared cursor until the
// list is drained. Keeps us well under GitHub's abuse thresholds on ~800 items.
async function runPool<T>(
  items: T[],
  concurrency: number,
  worker: (item: T, index: number) => Promise<void>,
): Promise<void> {
  let cursor = 0;
  async function work(): Promise<void> {
    while (cursor < items.length) {
      const index = cursor++;
      const item = items[index];
      if (item === undefined) break;
      await worker(item, index);
    }
  }
  await Promise.all(Array.from({length: Math.min(concurrency, items.length)}, work));
}

async function main() {
  const {repo, concurrency} = parseArgs(process.argv.slice(2));

  if (!(await fileExists(MANIFEST_PATH))) {
    throw new Error('No manifest found. Run `node scripts/harvest/list-backlog.ts` first.');
  }
  const manifest = JSON.parse(await readFile(MANIFEST_PATH, 'utf8')) as Manifest;
  const items: ManifestItem[] = manifest.items ?? [];
  console.log(`Filling queue from manifest: ${items.length} item(s), concurrency ${concurrency}\n`);

  let done = 0;
  let fetched = 0;
  let skipped = 0;
  const failures: Array<{number: number; error: string}> = [];

  await runPool(items, concurrency, async (item) => {
    try {
      const result = await harvestOne({number: item.number, repo, skipIfExists: true});
      if (result.skipped) skipped++;
      else fetched++;
    } catch (err) {
      failures.push({number: item.number, error: errorMessage(err)});
    }
    done++;
    if (done % 25 === 0 || done === items.length) {
      process.stdout.write(
        `  ${done}/${items.length}  (fetched ${fetched}, skipped ${skipped}, failed ${failures.length})\r`,
      );
    }
  });
  process.stdout.write('\n');

  // Verify completeness against the manifest, not against this run's counters —
  // a resumed run may have fetched nothing yet still leave the queue whole.
  const missing = [];
  for (const item of items) {
    if (!(await fileExists(recordPathFor(item.number)))) missing.push(item.number);
  }

  if (missing.length === 0) {
    manifest.harvestComplete = true;
    await writeFile(MANIFEST_PATH, `${JSON.stringify(manifest, null, 2)}\n`);
    console.log(
      `\nQueue full: ${items.length}/${items.length} records present. harvestComplete = true.`,
    );
    console.log(
      'Next: node scripts/harvest/status.ts   (then drain — see the harvest-triage skill)',
    );
  } else {
    console.log(`\nIncomplete: ${missing.length} record(s) still missing.`);
    if (failures.length) {
      console.log('Failures this run:');
      for (const f of failures.slice(0, 20)) console.log(`  ! #${f.number}: ${f.error}`);
      if (failures.length > 20) console.log(`  … and ${failures.length - 20} more`);
    }
    console.log('Rerun to retry — already-fetched records are skipped.');
    process.exitCode = 1;
  }
}

main().catch((err: unknown) => {
  console.error(`harvest-all failed: ${errorMessage(err)}`);
  process.exitCode = 1;
});
