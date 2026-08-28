/**
 * The emitted package, part by part.
 *
 * This is the panel that makes the rest of the page checkable: a reader who does not
 * believe what the lanes reported can open `xl/styles.xml` and read it.
 *
 * A preserved part gets its own mark. "The model did not interpret this and carried it
 * verbatim" is an honest state and a deliberate one, and leaving it unmarked would let it
 * look like a gap.
 */

import {computed, defineComponent, h, type PropType, ref, watch} from 'vue';

import {formatBytes} from '../playground/format.ts';
import {type PackageEntry, prettyXml} from '../playground/package.ts';

export default defineComponent({
  name: 'PartTree',
  props: {
    parts: {type: Array as PropType<readonly PackageEntry[]>, required: true},
    preserved: {type: Array as PropType<readonly string[]>, required: true},
  },
  setup(props) {
    const selectedPath = ref<string | undefined>(undefined);

    // A new package invalidates the selection: the same path in a different file is a
    // different part, and keeping the old one selected would show one file's bytes under
    // another file's heading.
    watch(
      () => props.parts,
      (parts) => {
        selectedPath.value = parts.find((part) => part.path.endsWith('workbook.xml'))?.path;
      },
      {immediate: true},
    );

    const selected = computed(() => props.parts.find((part) => part.path === selectedPath.value));

    return () => {
      const preserved = new Set(props.preserved);
      const part = selected.value;
      return h('div', {class: 'pg-parts'}, [
        h('div', {class: 'pg-parts__list'}, [
          h(
            'table',
            {class: 'pg-parts__table'},
            h(
              'tbody',
              props.parts.map((entry) =>
                h(
                  'tr',
                  {
                    key: entry.path,
                    class: entry.path === selectedPath.value ? 'is-open' : undefined,
                  },
                  [
                    h('td', [
                      h(
                        'button',
                        {
                          type: 'button',
                          class: 'pg-parts__path',
                          onClick: () => {
                            selectedPath.value = entry.path;
                          },
                        },
                        entry.path,
                      ),
                      preserved.has(entry.path)
                        ? h(
                            'span',
                            {
                              class: 'pg-badge',
                              title:
                                'Read but not interpreted. The model carries these bytes verbatim so a ' +
                                'round-trip keeps them.',
                            },
                            'preserved',
                          )
                        : null,
                    ]),
                    h('td', {class: 'pg-parts__size xlsx-figure'}, formatBytes(entry.byteLength)),
                  ],
                ),
              ),
            ),
          ),
        ]),
        h('div', {class: 'pg-parts__body'}, [
          part === undefined
            ? h('p', {class: 'pg-empty'}, 'Pick a part to read it.')
            : part.text === undefined
              ? h(
                  'p',
                  {class: 'pg-empty'},
                  `${part.path} is binary (${formatBytes(part.byteLength)}), so there is nothing to show as text.`,
                )
              : h('div', [
                  h(
                    'p',
                    {class: 'pg-note'},
                    'Indented here for reading. The emitted bytes carry no padding at all.',
                  ),
                  h('pre', {class: 'pg-xml'}, h('code', prettyXml(part.text))),
                ]),
        ]),
      ]);
    };
  },
});
