#!/usr/bin/env node
// No link in the built site may forget the base.
//
// The site is served from a subdirectory, so every internal URL has to carry `/ts-xlsx/`.
// VitePress adds it to the target of a *markdown* link and leaves raw HTML alone, which
// makes a hand-written `<a href="/docs/guide/">` a link that works on every developer's
// machine and 404s in production. `ignoreDeadLinks` cannot catch it, because it reads
// markdown rather than output.
//
// So this reads the output. Every `href` and `src` in the built HTML that starts with a
// single slash must start with the base, and one that does not fails the build. It is the
// only check here that needs `site:build` to have run, which is why it runs after it rather
// than in the invariants gate.
//
//   node www/scripts/check-built-links.ts

import {readdir, readFile} from 'node:fs/promises';
import {join, relative, resolve} from 'node:path';

import {repoRoot, siteBase} from './repo.ts';

const DIST = resolve(repoRoot, 'www', '.vitepress', 'dist');
const BASE = siteBase;
const ATTRIBUTE = /(?:href|src)="(\/[^"]*)"/g;

async function pages(dir: string): Promise<string[]> {
  const found: string[] = [];
  for (const entry of await readdir(dir, {withFileTypes: true})) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) found.push(...(await pages(path)));
    else if (entry.name.endsWith('.html')) found.push(path);
  }
  return found.sort();
}

const html = await pages(DIST).catch(() => {
  throw new Error(`${relative(repoRoot, DIST)} does not exist; run site:build first.`);
});

const offenders: string[] = [];
for (const page of html) {
  const source = await readFile(page, 'utf8');
  for (const match of source.matchAll(ATTRIBUTE)) {
    const url = match[1] ?? '';
    // A protocol-relative URL is not ours to prefix, and neither is the base itself.
    if (url.startsWith('//') || url.startsWith(BASE)) continue;
    offenders.push(`${relative(DIST, page).replaceAll('\\', '/')}  ${url}`);
  }
}

if (offenders.length > 0) {
  console.error(`check-built-links: ${offenders.length} link(s) missing the "${BASE}" base:\n`);
  for (const offender of offenders) console.error(`  ${offender}`);
  console.error(
    '\nA raw <a href="/..."> in markdown is the usual cause. VitePress adds the base to a ' +
      'markdown link and not to hand-written HTML; use the Door component, or write the ' +
      'link in markdown.',
  );
  process.exitCode = 1;
} else {
  console.log(`check-built-links: every internal link in ${html.length} pages carries "${BASE}"`);
}
