/**
 * What the site knows at build time and the browser cannot work out for itself.
 *
 * Two virtual modules, one plugin, because both need a syntax highlighter and standing up
 * two of those to serve one page would be silly.
 *
 * - `virtual:sample-sources` is each playground builder's own code. The page shows the code
 *   that produced the file it is showing, and the only way for that to stay true is to cut
 *   it out of `samples.ts` rather than keep a second copy.
 * - `virtual:site-facts` is the numbers the home page quotes, counted from the repository
 *   rather than typed into a paragraph. A hardcoded count is wrong the week after someone
 *   adds a corpus case, and wrong in the direction that flatters us.
 *
 * Highlighting here rather than in the browser means shiki never reaches the bundle, and
 * the markup that does reach it is this repository's own source, read off disk during the
 * build. The dual-theme output writes `--shiki-light` and `--shiki-dark` on each token,
 * which is the pair VitePress's own code blocks use, so a block follows the theme toggle
 * with no stylesheet of its own.
 */

import {readdirSync, readFileSync} from 'node:fs';
import {resolve} from 'node:path';

import {createHighlighter, type HighlighterGeneric} from 'shiki';

import {builderSource, SAMPLES} from '../playground/samples.ts';
import {readDocsSource} from './docs-source.ts';
import {pkg, repoRoot, siteBase} from './repo.ts';

const SAMPLE_SOURCES = 'virtual:sample-sources';
const SITE_FACTS = 'virtual:site-facts';
// The leading NUL is the convention that keeps other plugins and the resolver off a module
// that has no file behind it.
const RESOLVED = (id: string): string => `\u0000${id}`;

const SAMPLES_FILE = resolve(repoRoot, 'www', 'playground', 'samples.ts');
const QUICK_START_PAGE = resolve(repoRoot, 'docs', 'guide', 'README.md');

/**
 * The shape Vite needs, stated here rather than imported.
 *
 * `vite` is not a dependency of this repository: it arrives under `vitepress`, which pins
 * its own major. Declaring a `vite` dependency to name one type would mean pinning that
 * major a second time, in a second place, and living with the day the two disagree. Two
 * hooks are all this plugin has, and they are assignable to Vite's `Plugin` as written.
 */
interface VirtualModulePlugin {
  readonly name: string;
  readonly resolveId: (id: string) => string | undefined;
  readonly load: (id: string) => Promise<string | undefined>;
}

type Highlighter = HighlighterGeneric<'ts', 'github-light' | 'github-dark'>;

async function highlighter(): Promise<Highlighter> {
  return createHighlighter({themes: ['github-light', 'github-dark'], langs: ['ts']});
}

const asHtml = (shiki: Highlighter, code: string): string =>
  shiki.codeToHtml(code, {
    lang: 'ts',
    themes: {light: 'github-light', dark: 'github-dark'},
    defaultColor: false,
  });

/** How many `.md` files a directory holds, with an optional name to leave out. */
function countMarkdown(dir: string, except: readonly string[] = []): number {
  return readdirSync(resolve(repoRoot, dir)).filter(
    (name) => name.endsWith('.md') && !except.includes(name),
  ).length;
}

/**
 * The quick start, taken from the guide page whose samples are executed on every verify.
 *
 * Copying five lines onto the home page would be cheap and would be wrong within a release:
 * the copy is the one nothing runs. Taking the guide's first block means the home page's
 * code is code a gate already proved works.
 */
function quickStart(): string {
  const page = readFileSync(QUICK_START_PAGE, 'utf8');
  const match = /```ts\r?\n([\s\S]*?)```/.exec(page);
  if (match?.[1] === undefined) {
    throw new Error(
      `${QUICK_START_PAGE} has no \`\`\`ts block, so the home page has no quick start to show. ` +
        'It is taken from there rather than copied so that the code on the home page is code ' +
        'check-samples has run.',
    );
  }
  return match[1].trimEnd();
}

function facts(): Record<string, string | number> {
  const dependencies = Object.keys(pkg.dependencies);
  return {
    // The one thing here the browser needs rather than the reader: a component that builds a
    // url has to prefix it, and this is where the answer already lives.
    base: siteBase,
    runtimeDependencies: dependencies.length,
    runtimeDependencyNames: dependencies.join(', '),
    corpusCases: readdirSync(resolve(repoRoot, 'test', 'corpus', 'cases')).filter((name) =>
      name.endsWith('.case.ts'),
    ).length,
    decisionRecords: countMarkdown('docs/decisions', ['README.md']),
    specNotes: countMarkdown('docs/knowledge/specs', ['README.md']),
    guidePages: countMarkdown('docs/guide'),
    docsPages: readDocsSource().pages.length,
  };
}

/**
 * Note for whoever edits `samples.ts` or a counted tree with the dev server running: these
 * modules are built once per server start, so a change needs a restart to be picked up. Every
 * `site:build` rebuilds them, so nothing published can be stale. Fixing it properly means
 * holding a Vite module graph open, which is more machinery than a restart is worth.
 */
export function virtualModules(): VirtualModulePlugin {
  return {
    name: 'ts-xlsx:virtual-modules',
    resolveId: (id) => (id === SAMPLE_SOURCES || id === SITE_FACTS ? RESOLVED(id) : undefined),
    load: async (id) => {
      if (id === RESOLVED(SAMPLE_SOURCES)) {
        const shiki = await highlighter();
        const moduleText = readFileSync(SAMPLES_FILE, 'utf8');
        const byId: Record<string, string> = {};
        for (const sample of SAMPLES) {
          byId[sample.id] = asHtml(shiki, builderSource(moduleText, sample.id));
        }
        shiki.dispose();
        return `export default ${JSON.stringify(byId)};`;
      }
      if (id === RESOLVED(SITE_FACTS)) {
        const shiki = await highlighter();
        const value = {...facts(), quickStartHtml: asHtml(shiki, quickStart())};
        shiki.dispose();
        return `export default ${JSON.stringify(value)};`;
      }
      return undefined;
    },
  };
}
