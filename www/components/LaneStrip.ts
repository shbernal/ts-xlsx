/**
 * One box per lane, reporting what it produced and how long it took.
 *
 * A failed lane is content, not an absence. The library's errors are a documented public
 * surface, so the box shows the message verbatim; a demo that folded them into "something
 * went wrong" would be throwing away the best thing it had to show. The lanes after a
 * failure say they did not run, rather than sitting empty and looking broken too.
 */

import {defineComponent, h, type PropType} from 'vue';

import {formatMillis} from '../playground/format.ts';

export type LaneStatus = 'ok' | 'failed' | 'skipped';

export interface LaneView {
  /** What a consumer would call to do this: `writeXlsx`. */
  readonly call: string;
  readonly title: string;
  readonly status: LaneStatus;
  /** One line of what the lane produced, or why it did not run. */
  readonly detail: string;
  readonly elapsedMs: number | undefined;
  /** The library's own message, verbatim, when the lane failed. */
  readonly error: string | undefined;
}

const MARK: Record<LaneStatus, string> = {ok: 'ok', failed: 'failed', skipped: 'not run'};

export default defineComponent({
  name: 'LaneStrip',
  props: {
    lanes: {type: Array as PropType<readonly LaneView[]>, required: true},
  },
  setup(props) {
    return () =>
      h(
        'div',
        {class: 'pg-lanes'},
        props.lanes.map((lane) =>
          h('div', {class: ['pg-lane', `pg-lane--${lane.status}`], key: lane.call}, [
            h('div', {class: 'pg-lane__head'}, [
              h('code', {class: 'pg-lane__call'}, lane.call),
              h('span', {class: 'pg-lane__status'}, MARK[lane.status]),
            ]),
            h('div', {class: 'pg-lane__title'}, lane.title),
            h('div', {class: 'pg-lane__detail'}, lane.detail),
            lane.error === undefined ? null : h('pre', {class: 'pg-lane__error'}, lane.error),
            lane.elapsedMs === undefined
              ? null
              : h('div', {class: 'pg-lane__time xlsx-figure'}, formatMillis(lane.elapsedMs)),
          ]),
        ),
      );
  },
});
