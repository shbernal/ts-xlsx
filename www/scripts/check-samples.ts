#!/usr/bin/env node
// Every code sample in the guide is run.
//
// Prose can be rewritten cheaply. A sample that no longer compiles is how documentation
// stops being trusted, and it fails silently: nobody notices until a reader copies it. So
// each fenced `ts` block under `docs/guide/` is extracted, pointed at this repository's own
// `src/`, and executed. A block that throws fails the check, naming the page and the line.
//
// Only `docs/guide/`. A block in an ADR is a record of what was decided on a day, and some
// deliberately show code that no longer compiles; running those would make the record
// unwritable. The guide is the tree whose blocks are instructions to a reader.
//
// One process runs every block, each as its own module, so the library is type-stripped and
// loaded once rather than once per sample. That is the difference between a check that runs
// on every verify and one that gets skipped.
//
// A block that is illustration rather than instruction opts out with a marker comment on the
// line above its fence. The marker is the only way out, so opting out shows up in a diff.
//
//   node www/scripts/check-samples.ts

import {mkdir, readdir, readFile, rm, writeFile} from 'node:fs/promises';
import {join, relative, resolve} from 'node:path';
import {pathToFileURL} from 'node:url';

import {repoRoot} from './repo.ts';

const GUIDE = resolve(repoRoot, 'docs', 'guide');
const SCRATCH = resolve(repoRoot, '.tmp', 'guide-samples');
const PACKAGE = '@shbernal/ts-xlsx';

/** Sits above a fence, with an optional reason after it. */
const OPT_OUT = /^<!--\s*sample:\s*illustrative\b/;
const FENCE_OPEN = /^```ts\s*$/;
const FENCE_CLOSE = /^```\s*$/;

interface Sample {
  /** `docs/guide/writing.md:42`, which is what a failure has to name. */
  readonly where: string;
  readonly file: string;
  readonly code: string;
}

async function guidePages(dir: string): Promise<string[]> {
  const found: string[] = [];
  for (const entry of await readdir(dir, {withFileTypes: true})) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) found.push(...(await guidePages(path)));
    else if (entry.name.endsWith('.md')) found.push(path);
  }
  return found.sort();
}

function extract(source: string, page: string): Sample[] {
  const lines = source.split('\n');
  const samples: Sample[] = [];
  let index = 0;
  let ordinal = 0;
  while (index < lines.length) {
    if (!FENCE_OPEN.test(lines[index] ?? '')) {
      index += 1;
      continue;
    }
    const fenceLine = index + 1;
    index += 1;
    const body: string[] = [];
    while (index < lines.length && !FENCE_CLOSE.test(lines[index] ?? '')) {
      body.push(lines[index] ?? '');
      index += 1;
    }
    index += 1;
    ordinal += 1;
    // The marker sits above the fence, across a blank line if the author left one.
    const previous = (lines[fenceLine - 2] ?? '').trim();
    const marker = previous === '' ? (lines[fenceLine - 3] ?? '').trim() : previous;
    if (OPT_OUT.test(marker)) continue;
    samples.push({
      where: `${page}:${fenceLine}`,
      file: `${page.replaceAll('/', '-').replace(/\.md$/, '')}-${ordinal}.ts`,
      code: body.join('\n'),
    });
  }
  return samples;
}

/**
 * The specifier a reader writes, pointed at the tree this repository actually has.
 *
 * The guide has to show `@shbernal/ts-xlsx`, because that is what a reader types. Running
 * that against the published package would test whatever version happens to be installed
 * rather than this commit, so the specifier is rewritten to the source barrel its subpath
 * maps to.
 */
function localise(code: string, from: string): string {
  const barrel = (subpath: string): string => {
    const target =
      subpath === ''
        ? resolve(repoRoot, 'src', 'index.ts')
        : resolve(repoRoot, 'src', 'entries', `${subpath}.ts`);
    return relative(from, target).replaceAll('\\', '/');
  };
  return code.replace(
    new RegExp(`(['"])${PACKAGE}(/[a-z]+)?\\1`, 'g'),
    (_whole, quote: string, subpath: string | undefined) =>
      `${quote}${barrel(subpath === undefined ? '' : subpath.slice(1))}${quote}`,
  );
}

const pages = await guidePages(GUIDE);
const samples: Sample[] = [];
for (const page of pages) {
  const relativePath = relative(repoRoot, page).replaceAll('\\', '/');
  samples.push(...extract(await readFile(page, 'utf8'), relativePath));
}

await rm(SCRATCH, {recursive: true, force: true});
await mkdir(SCRATCH, {recursive: true});
for (const sample of samples) {
  await writeFile(resolve(SCRATCH, sample.file), `${localise(sample.code, SCRATCH)}\n`);
}

// Output is captured rather than printed: a guide is full of `console.log`, and forty of
// them would bury the one line that matters. Capturing it also makes the guide's own idiom
// checkable. A sample that states a claim writes it as `console.log(a === b); // true`, so a
// logged line that reads exactly `false` is a claim the library no longer honours, and this
// treats it as a failure. Without that, a sample could run perfectly and print the opposite
// of what the paragraph above it says.
const real = {log: console.log, error: console.error, warn: console.warn};
const failures: {readonly where: string; readonly error: unknown}[] = [];
for (const sample of samples) {
  const printed: string[] = [];
  const capture = (...args: readonly unknown[]): void => {
    printed.push(args.map((arg) => (typeof arg === 'string' ? arg : String(arg))).join(' '));
  };
  Object.assign(console, {log: capture, error: capture, warn: capture});
  try {
    await import(pathToFileURL(resolve(SCRATCH, sample.file)).href);
    if (printed.includes('false')) {
      throw new Error(
        `a claim in this sample printed false. What it printed:\n${printed.join('\n')}`,
      );
    }
  } catch (err: unknown) {
    failures.push({where: sample.where, error: err});
  } finally {
    Object.assign(console, real);
  }
}

if (failures.length > 0) {
  for (const failure of failures) {
    const {error} = failure;
    console.error(`\n─── ${failure.where} ───`);
    console.error(error instanceof Error ? (error.stack ?? error.message) : String(error));
  }
  console.error(`\ncheck-samples: ${failures.length} of ${samples.length} samples failed`);
  process.exitCode = 1;
} else {
  console.log(
    `check-samples: ${samples.length} runnable samples across ${pages.length} guide pages`,
  );
}
