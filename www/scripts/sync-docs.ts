#!/usr/bin/env node
// Writes www/docs/, the generated mirror of the tracked docs/ tree.
//
// The mirror is removed and rewritten whole on every run rather than diffed. A stale page
// in a generated tree is invisible until someone reads it, which is the one failure this
// whole arrangement exists to prevent.
//
// Each mirrored file keeps the source's own filename, README.md included; the site serves a
// directory's README at the directory root through VitePress's `rewrites` instead. That is
// what lets the theme's edit link point at the tracked file with no mapping of its own: the
// mirror path below srcDir and the repository path are the same string.
//
//   node www/scripts/sync-docs.ts

import {mkdir, rm, writeFile} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';

import {type DocPage, readDocsSource} from './docs-source.ts';
import {repoRoot} from './repo.ts';

const MIRROR = resolve(repoRoot, 'www', 'docs');

/** JSON strings are valid YAML double-quoted scalars, and a title here holds colons and quotes. */
const yaml = (key: string, value: string): string => `${key}: ${JSON.stringify(value)}\n`;

function render(page: DocPage): string {
  let frontmatter = '---\n';
  frontmatter += yaml('title', page.title);
  if (page.description !== undefined) frontmatter += yaml('description', page.description);
  frontmatter += '---\n\n';
  return (
    `${frontmatter}<!-- Generated from docs/${page.sourcePath}. ` +
    `Edit that file: this copy is rewritten on every build. -->\n\n${page.body}`
  );
}

const source = readDocsSource();

await rm(MIRROR, {recursive: true, force: true});
await Promise.all(
  source.pages.map(async (page) => {
    const target = resolve(MIRROR, page.sourcePath);
    await mkdir(dirname(target), {recursive: true});
    await writeFile(target, render(page));
  }),
);

console.log(`www/docs: ${source.pages.length} pages written from docs/`);
