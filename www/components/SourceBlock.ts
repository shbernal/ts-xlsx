/**
 * The code that produced the file on screen.
 *
 * The markup is generated at build time, by the `sample-sources` plugin in the site config,
 * out of the same module the builders live in. It is written with `v-html` because it is
 * highlighted markup rather than text, and it is safe to do so for one reason worth stating:
 * the only input is this repository's own source, read off disk during the build. Nothing a
 * reader supplies reaches this component.
 */

import {defineComponent, h, type PropType} from 'vue';

export default defineComponent({
  name: 'SourceBlock',
  props: {
    html: {type: String as PropType<string | undefined>, default: undefined},
    caption: {type: String, required: true},
  },
  setup(props) {
    return () =>
      h('div', {class: 'pg-source'}, [
        h('p', {class: 'pg-note'}, props.caption),
        props.html === undefined
          ? h(
              'p',
              {class: 'pg-empty'},
              'A file you dropped has no source here. This is the code behind a built-in sample.',
            )
          : h('div', {class: 'pg-source__code', innerHTML: props.html}),
      ]);
  },
});
