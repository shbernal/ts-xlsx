/**
 * Repository facts the site needs, read from the file that already owns them.
 *
 * Nothing here may be typed a second time in `config.ts` or in a markdown page: the
 * repository URL and the package description live in `package.json`, and a copy of
 * either would be right on the day it was written and wrong later.
 */

import {readFileSync} from 'node:fs';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

/** Resolved from this file, not from `process.cwd()`, so the scripts run from anywhere. */
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');

interface PackageJson {
  readonly name: string;
  readonly description: string;
  readonly author: string;
  readonly license: string;
  readonly repository: {readonly url: string};
}

export const pkg = JSON.parse(
  readFileSync(resolve(repoRoot, 'package.json'), 'utf8'),
) as PackageJson;

/** `https://github.com/owner/repo.git` -> `https://github.com/owner/repo` */
export const repoUrl = pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '');

/** The default branch. Ours is `master`; every blob and edit link has to say so. */
export const branch = 'master';

/** A link to a file that lives in the repository and is deliberately not published here. */
export function blobUrl(path: string): string {
  return `${repoUrl}/blob/${branch}/${path}`;
}
