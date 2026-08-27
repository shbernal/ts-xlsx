import {defineConfig} from 'vitepress';

import {blobUrl, branch, pkg, repoUrl} from '../scripts/repo.ts';

export default defineConfig({
  // GitHub Pages serves the repo at https://shbernal.github.io/ts-xlsx/. Every URL on the
  // site goes through this, so nothing may hardcode a leading `/`. The override exists so a
  // deploy preview can be served from a different prefix without editing the config.
  base: process.env['VITEPRESS_BASE'] ?? '/ts-xlsx/',
  title: 'ts-xlsx',
  description: pkg.description,
  lang: 'en-GB',
  cleanUrls: true,

  // Without this, VitePress publishes every markdown file under the root, and this
  // directory's README is a note to whoever works on the site, not a page for a reader.
  srcExclude: ['README.md'],

  // Left on deliberately. A docs page written for the repository tree links to its
  // neighbours by relative path, and the mirror that brings those pages into the site has to
  // rewrite every one of them; this setting is what fails the build when it misses one.
  ignoreDeadLinks: false,

  themeConfig: {
    nav: [
      // The changelog stays a repository file: it is written for someone standing at a
      // version, not for someone browsing a site, and publishing it would make the site a
      // second place it has to be right.
      {text: 'Changelog', link: blobUrl('CHANGELOG.md')},
      {text: 'GitHub', link: repoUrl},
    ],

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

  vite: {
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
