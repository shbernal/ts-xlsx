/**
 * A link out of the home page, to somewhere worth going.
 *
 * A component rather than an `<a>` written into the markdown, for one reason that cost an
 * afternoon to find: VitePress rewrites the target of a *markdown* link to include the site's
 * base and leaves raw HTML alone. The site is served from a subdirectory, so a hand-written
 * `<a href="/docs/guide/">` is a working link on every developer's machine and a 404 in
 * production, and `ignoreDeadLinks` never sees it because it reads markdown rather than
 * output.
 *
 * Prefixing here means the markdown cannot forget it, and the base comes from the same
 * constant the config sets VitePress's own `base` from. `www/scripts/check-built-links.ts`
 * is the gate that catches the next link written by hand anyway.
 */

import FACTS from 'virtual:site-facts';
import {defineComponent, h} from 'vue';

const ABSOLUTE = /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i;

const withBase = (href: string): string =>
  ABSOLUTE.test(href) ? href : `${FACTS.base.replace(/\/$/, '')}${href}`;

export default defineComponent({
  name: 'Door',
  props: {
    href: {type: String, required: true},
    title: {type: String, required: true},
  },
  setup(props, {slots}) {
    return () =>
      h('a', {class: 'home-door', href: withBase(props.href)}, [
        h('strong', props.title),
        h('span', slots['default']?.()),
      ]);
  },
});
