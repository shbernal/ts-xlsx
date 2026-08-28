/**
 * `virtual:sample-sources`: each sample builder's own code, highlighted at build time.
 *
 * The page shows the code that produced the file it is showing, and the only way for that
 * to stay true is to cut it out of `samples.ts` rather than keep a second copy. Doing it
 * here rather than in the browser means shiki never reaches the bundle, and the markup that
 * does reach it is this repository's own source, read off disk during the build.
 *
 * The dual-theme output writes `--shiki-light` and `--shiki-dark` on each token, which is
 * the same pair VitePress's own code blocks use, so the block follows the theme toggle with
 * no stylesheet of its own.
 */

import {readFileSync} from 'node:fs';
import {resolve} from 'node:path';

import {createHighlighter} from 'shiki';

import {builderSource, SAMPLES} from '../playground/samples.ts';
import {repoRoot} from './repo.ts';

const MODULE_ID = 'virtual:sample-sources';
// The leading NUL is the convention that keeps other plugins and the resolver off a module
// that has no file behind it.
const RESOLVED_ID = `\u0000${MODULE_ID}`;
const SOURCE_FILE = resolve(repoRoot, 'www', 'playground', 'samples.ts');

/**
 * The shape Vite needs, stated here rather than imported.
 *
 * `vite` is not a dependency of this repository: it arrives under `vitepress`, which pins
 * its own major. Declaring a `vite` dependency to name one type would mean pinning that
 * major a second time, in a second place, and living with the day the two disagree. Three
 * hooks are all this plugin has, and they are assignable to Vite's `Plugin` as written.
 */
interface VirtualModulePlugin {
  readonly name: string;
  readonly resolveId: (id: string) => string | undefined;
  readonly load: (id: string) => Promise<string | undefined>;
}

/**
 * Note for whoever edits `samples.ts` with the dev server running: this module is built once
 * per server start, so a changed builder needs a restart to be re-highlighted. The page
 * itself hot-reloads, and every `site:build` re-reads the file, so nothing published can be
 * stale. Fixing it properly means holding a Vite module graph open, which is more machinery
 * than a restart is worth.
 */
export function sampleSources(): VirtualModulePlugin {
  return {
    name: 'ts-xlsx:sample-sources',
    resolveId: (id) => (id === MODULE_ID ? RESOLVED_ID : undefined),
    load: async (id) => {
      if (id !== RESOLVED_ID) return undefined;
      const moduleText = readFileSync(SOURCE_FILE, 'utf8');
      const highlighter = await createHighlighter({
        themes: ['github-light', 'github-dark'],
        langs: ['ts'],
      });
      const byId: Record<string, string> = {};
      for (const sample of SAMPLES) {
        byId[sample.id] = highlighter.codeToHtml(builderSource(moduleText, sample.id), {
          lang: 'ts',
          themes: {light: 'github-light', dark: 'github-dark'},
          defaultColor: false,
        });
      }
      highlighter.dispose();
      return `export default ${JSON.stringify(byId)};`;
    },
  };
}
