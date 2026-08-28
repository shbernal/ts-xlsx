import {defineConfig} from 'vitepress';

import {readDocsSource} from '../scripts/docs-source.ts';
import {blobUrl, branch, pkg, repoUrl, siteBase} from '../scripts/repo.ts';
import {virtualModules} from '../scripts/virtual-modules.ts';

// Read here rather than restated: the sidebar is the manifest, so a page added to one and
// not the other fails the config instead of quietly failing a reader. This throws on any
// drift, which is why `site:check` can be a sub-second gate and the build needs no second
// opinion.
const docs = readDocsSource();

const firstPage = docs.groups[0]?.pages[0]?.route ?? '/';

export default defineConfig({
  // GitHub Pages serves the repo at https://shbernal.github.io/ts-xlsx/. Every URL on the
  // site goes through this, so nothing may hardcode a leading `/`.
  base: siteBase,
  title: 'ts-xlsx',
  description: pkg.description,
  lang: 'en-GB',
  cleanUrls: true,

  // Without this, VitePress publishes every markdown file under the root, and this
  // directory's README is a note to whoever works on the site, not a page for a reader.
  srcExclude: ['README.md'],

  // A directory's README is served at the directory root. The mirror keeps the source
  // filename so that `editLink`'s `:path`, which is a page's own file path, is already the
  // path of the tracked file it was generated from.
  rewrites: docs.rewrites,

  // Left on deliberately. A docs page written for the repository tree links to its
  // neighbours by relative path, and the mirror that brings those pages into the site has to
  // rewrite every one of them; this setting is what fails the build when it misses one.
  ignoreDeadLinks: false,

  themeConfig: {
    nav: [
      // The changelog stays a repository file: it is written for someone standing at a
      // version, not for someone browsing a site, and publishing it would make the site a
      // second place it has to be right.
      // `firstPage` rather than a typed route: the guide's first page is decided by the
      // manifest, and a nav entry that named it separately would be the copy that rots.
      {text: 'Guide', link: firstPage},
      {text: 'API', link: '/docs/api/'},
      {text: 'Playground', link: '/playground'},
      {text: 'Changelog', link: blobUrl('CHANGELOG.md')},
      {text: 'GitHub', link: repoUrl},
    ],

    sidebar: {
      '/docs/': docs.groups.map((group) => ({
        text: group.text,
        collapsed: group.collapsed,
        items: group.pages.map((page) => ({text: page.title, link: page.route})),
      })),
    },

    search: {provider: 'local'},
    socialLinks: [{icon: 'github', link: repoUrl}],

    editLink: {
      pattern: `${repoUrl}/edit/${branch}/:path`,
      text: 'Edit this page on GitHub',
    },

    footer: {
      message: `Released under the ${pkg.license} License.`,
      copyright: `Copyright (c) ${pkg.author}`,
    },
  },

  markdown: {
    config: (md) => {
      // Inline code is literal text, and without this Vue interpolates a `{{ ... }}` inside
      // it as an expression. The spec note on template placeholders is written almost
      // entirely in mustache syntax and failed the build for it, naming a column that does
      // not exist on the line it named. Fenced blocks already come out with `v-pre`; this
      // gives an inline span the same guarantee, which is the one it always meant to have.
      const renderCodeInline = md.renderer.rules.code_inline;
      md.renderer.rules.code_inline = (tokens, index, options, env, self) => {
        tokens[index]?.attrSet('v-pre', '');
        return renderCodeInline === undefined
          ? self.renderToken(tokens, index, options)
          : renderCodeInline(tokens, index, options, env, self);
      };
    },
  },

  vite: {
    plugins: [virtualModules()],
    build: {
      // The default 500 kB warning is aimed at a chunk that blocks first paint. The one
      // chunk here that passes it is the local search index over 247 pages, 1.5 MB, fetched
      // only when a reader opens search. Raised rather than switched off, and only to just
      // above what that index costs, so the next chunk to cross the line still says so.
      chunkSizeWarningLimit: 2000,
    },
    resolve: {
      // This repo declares `vue` itself and VitePress carries its own copy. pnpm resolves
      // those to two versions, and two Vue runtimes in one page means `provide`/`inject` and
      // the app instance stop matching across the boundary, which shows up as a component
      // that mounts but sees none of the theme's context. Naming it here collapses both
      // specifiers onto one copy.
      dedupe: ['vue'],
    },
  },
});
