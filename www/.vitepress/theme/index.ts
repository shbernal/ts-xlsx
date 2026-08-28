/**
 * The default theme, extended in one place.
 *
 * Extended and not replaced: the documentation is the bulk of this site, and the default
 * layout is what makes the whole docs tree navigable. What this file is for is the stylesheet and the
 * components a page can mount. Restyling belongs in CSS, never here.
 *
 * The playground is registered asynchronously, and that is the whole reason this file has a
 * function in it. A static import would put the library, the zip codec and the grid model
 * into the chunk every page of the site loads, to serve one page; and it would break the
 * build outright, because VitePress pre-renders every page in Node while the component wants
 * a DOM. Fetched when it mounts, and it only mounts inside `<ClientOnly>`.
 */

import FACTS from 'virtual:site-facts';
import type {Theme} from 'vitepress';
import DefaultTheme from 'vitepress/theme';
import {defineAsyncComponent, h} from 'vue';

import Door from '../../components/Door.ts';
import QuickStart from '../../components/QuickStart.ts';

import './style.css';

export default {
  extends: DefaultTheme,
  enhanceApp({app}) {
    // Counted from the repository at build time and reachable from any page as `$facts`, so
    // a number in a paragraph is read rather than typed. A hardcoded corpus count is wrong
    // the week after someone adds a case, and wrong in the direction that flatters us.
    // `globalProperties` is typed as an index signature over `any`, so the assignment has to
    // be widened deliberately rather than silently: what a page reads as `$facts` is checked
    // by the `SiteFacts` declaration in `www/env.d.ts`, not here.
    (app.config.globalProperties as Record<string, unknown>)['$facts'] = FACTS;

    // Static, because it is markup with no state and belongs in the pre-rendered HTML of the
    // page most visitors see first.
    app.component('QuickStart', QuickStart);
    app.component('Door', Door);

    app.component(
      'Playground',
      defineAsyncComponent({
        loader: async () => import('../../components/Playground.ts'),
        loadingComponent: {
          render: () => h('p', {class: 'pg-booting'}, 'Loading ts-xlsx into this tab...'),
        },
        errorComponent: {
          render: () =>
            h(
              'p',
              {class: 'pg-booting'},
              'The playground failed to load, so this space is empty. Nothing on this page ' +
                'stands in for a workbook it did not write.',
            ),
        },
      }),
    );
  },
} satisfies Theme;
