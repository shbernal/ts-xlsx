// CLAUDE.md and AGENTS.md must stay byte-identical.
//
// They were one file behind a symlink, so every runtime converged on the same rules by
// construction. That construction does not survive Windows: creating a symlink needs
// admin (or Developer Mode), so a checkout there materializes AGENTS.md as a ~9-byte
// text file containing the literal string "CLAUDE.md" — and an AGENTS.md-reading agent
// silently gets no constitution at all. The failure is silent, which is the worst kind.
//
// So AGENTS.md is now a real file, and this check is what replaces the symlink's
// guarantee: duplication is only safe when drift is impossible, and the way we make
// drift impossible here is a machine check, not a promise to remember.
//
// Comparison is line-ending-insensitive on purpose — the two files are the same
// document, and a checkout that hands one CRLF is not a divergence of content.

import {readFileSync} from 'node:fs';

const SOURCE = 'CLAUDE.md';
const MIRROR = 'AGENTS.md';

const normalize = (text: string) => text.replace(/\r\n/g, '\n');

const source = normalize(readFileSync(SOURCE, 'utf8'));
const mirror = normalize(readFileSync(MIRROR, 'utf8'));

if (source === mirror) {
  console.log(`constitution: ${MIRROR} matches ${SOURCE} (${source.length} chars)`);
} else {
  const sourceLines = source.split('\n');
  const mirrorLines = mirror.split('\n');
  const at = sourceLines.findIndex((line, i) => line !== mirrorLines[i]);
  console.error(
    `\n${MIRROR} has drifted from ${SOURCE}.\n\n` +
      (at === -1
        ? `  ${SOURCE} has ${sourceLines.length} line(s), ${MIRROR} has ${mirrorLines.length}.\n`
        : `  First difference at line ${at + 1}:\n` +
          `    ${SOURCE}: ${JSON.stringify(sourceLines[at])}\n` +
          `    ${MIRROR}: ${JSON.stringify(mirrorLines[at] ?? '<missing>')}\n`) +
      `\nThese two are one document. Edit ${SOURCE}, then copy it over ${MIRROR}.\n`,
  );
  process.exit(1);
}
