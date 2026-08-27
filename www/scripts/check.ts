#!/usr/bin/env node
// The drift gate between docs/ and the site.
//
// It reads and validates the docs tree without writing anything, so it can run on every
// verify: the full site build takes seconds and pulls in Vite, while this is the part that
// actually catches drift. Everything it can fail on is listed in www/scripts/docs-source.ts,
// and each one is a throw, never a warning.
//
//   node www/scripts/check.ts

import {type DocsSource, readDocsSource} from './docs-source.ts';

let source: DocsSource;
try {
  source = readDocsSource();
} catch (err: unknown) {
  // One legible line. Every throw in docs-source.ts names a file, a line and what to do
  // about it, and a stack trace through a walk of 247 pages buries all three.
  console.error(`site:check: ${err instanceof Error ? err.message : String(err)}`);
  process.exit(1);
}

const withoutDescription = source.pages.filter((page) => page.description === undefined);

console.log(`docs: ${source.pages.length} pages in ${source.groups.length} groups`);
for (const group of source.groups) {
  console.log(`  ${group.text.padEnd(14)} ${String(group.pages.length).padStart(3)}`);
}

// Reported, not failed. A page whose opening block is a code fence has no sentence to take
// a description from, and the honest answer is to publish none. Printing the list is what
// keeps the number from creeping: a derivation that quietly stops working on half the tree
// looks exactly like one that works.
if (withoutDescription.length > 0) {
  console.log(`\nno lead paragraph, so no description (${withoutDescription.length}):`);
  for (const page of withoutDescription) console.log(`  docs/${page.sourcePath}`);
}
