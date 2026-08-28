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
export const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');

interface PackageJson {
  readonly name: string;
  readonly description: string;
  readonly author: string;
  readonly license: string;
  readonly dependencies: Readonly<Record<string, string>>;
  readonly repository: {readonly url: string};
}

export const pkg = JSON.parse(
  readFileSync(resolve(repoRoot, 'package.json'), 'utf8'),
) as PackageJson;

/** `https://github.com/owner/repo.git` -> `https://github.com/owner/repo` */
export const repoUrl = pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '');

/** The default branch. Ours is `master`; every blob and edit link has to say so. */
export const branch = 'master';

/**
 * Where the site is served from, defined once.
 *
 * GitHub Pages puts the repository under a subdirectory, so every internal url on the site
 * carries this. The config sets VitePress's `base` from it, the facts module hands it to the
 * one component that builds a url itself, and the post-build link check asserts against it.
 * Three readers, one definition, and an override so a preview deploy can be served from a
 * different prefix without editing anything.
 */
export const siteBase = process.env['VITEPRESS_BASE'] ?? '/ts-xlsx/';

/** A link to a file that lives in the repository and is deliberately not published here. */
export function blobUrl(path: string): string {
  return `${repoUrl}/blob/${branch}/${path}`;
}

/** The same, for a directory. GitHub serves those under `/tree/`, and a blob url only redirects. */
export function treeUrl(path: string): string {
  return `${repoUrl}/tree/${branch}/${path.replace(/\/$/, '')}`;
}
