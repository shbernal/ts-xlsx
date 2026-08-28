/**
 * The quick start, which is the guide's first sample rather than a copy of it.
 *
 * Copying five lines onto the home page would be cheap and would be wrong within a release,
 * because the copy is the one nothing runs. This block is cut out of `docs/guide/README.md`
 * at build time and highlighted there, so the code on the front page is code
 * `check-samples` has already executed.
 *
 * `innerHTML` for the same reason `SourceBlock` uses it: the only input is this repository's
 * own source, read off disk during the build, and nothing a reader supplies reaches here.
 */

import FACTS from 'virtual:site-facts';
import {defineComponent, h} from 'vue';

export default defineComponent({
  name: 'QuickStart',
  setup() {
    return () =>
      h('div', {class: 'pg-source__code home-quickstart', innerHTML: FACTS.quickStartHtml});
  },
});
