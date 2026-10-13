#!/usr/bin/env node
// Derive the version used by the `next` release channel from the stable manifest version.
// The generated version is written only into the workflow checkout. The repository and its tags
// continue to name stable releases.

import {resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseArgs} from 'node:util';

import {readPackageJson} from './repo.ts';

const SEMVER = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z.-]+))?$/;
const SHA_LENGTH = 7;

/** Build `<base>-next.<run>.g<sha>` for one `main` commit. */
export function nextVersion(version: string, run: number, sha: string): string {
  const match = SEMVER.exec(version);
  if (match === null) throw new Error(`package version ${JSON.stringify(version)} is not SemVer`);
  if (!Number.isSafeInteger(run) || run < 1) {
    throw new Error(`run number must be a positive integer, got ${run}`);
  }
  if (!/^[0-9a-f]{7,40}$/.test(sha)) {
    throw new Error(
      `commit must be a 7 to 40 character lowercase hex hash, got ${JSON.stringify(sha)}`,
    );
  }

  const [, major, minor, patch, prerelease] = match;
  const base =
    prerelease === undefined
      ? `${major}.${minor}.${Number(patch) + 1}`
      : `${major}.${minor}.${patch}`;
  return `${base}-next.${run}.g${sha.slice(0, SHA_LENGTH)}`;
}

function main(): void {
  const {values} = parseArgs({
    options: {
      run: {type: 'string'},
      sha: {type: 'string'},
    },
    strict: true,
  });
  if (values.run === undefined || values.sha === undefined) {
    throw new Error('usage: node scripts/next-version.ts --run <number> --sha <commit>');
  }
  console.log(nextVersion(readPackageJson().version, Number(values.run), values.sha));
}

if (process.argv[1] !== undefined && fileURLToPath(import.meta.url) === resolve(process.argv[1]))
  main();
