/**
 * The playground: pick a sample or drop a file, and watch the library run on it.
 *
 * Every decision that could be wrong lives under `www/playground/`, where a test can reach
 * it. What is left here is assignment and markup: which lane ran, what it produced, and
 * which panel is open. That split is the point of writing components as render functions
 * rather than templates, since this file is read by `tsc`, oxlint and oxfmt and a template
 * would be read by none of them.
 *
 * The bytes never leave the tab. That is a property of this code, not a promise: there is no
 * `fetch`, no form, and no url anywhere in the modules this page runs.
 */

import SAMPLE_SOURCES from 'virtual:sample-sources';
import {computed, defineComponent, h, onBeforeUnmount, ref, shallowRef} from 'vue';

import type {Workbook} from '../../src/index.ts';
import {formatBytes, formatCount} from '../playground/format.ts';
import {DEFAULT_BOUNDS, type Grid, toGrid} from '../playground/grid.ts';
import {type Lane, read, roundTrip, write} from '../playground/lanes.ts';
import {listParts, type PackageEntry} from '../playground/package.ts';
import {findSample, SAMPLES} from '../playground/samples.ts';
import LaneStrip, {type LaneView} from './LaneStrip.ts';
import PartTree from './PartTree.ts';
import SheetGrid from './SheetGrid.ts';
import SourceBlock from './SourceBlock.ts';

type Panel = 'sheet' | 'parts' | 'source';

/** What a run produced, or how far it got. */
interface RunState {
  readonly label: string;
  readonly sampleId: string | undefined;
  readonly lanes: readonly LaneView[];
  readonly bytes: Uint8Array | undefined;
  readonly grid: Grid | undefined;
  readonly sheetNames: readonly string[];
  readonly parts: readonly PackageEntry[];
  readonly preserved: readonly string[];
}

const EMPTY: RunState = {
  label: '',
  sampleId: undefined,
  lanes: [],
  bytes: undefined,
  grid: undefined,
  sheetNames: [],
  parts: [],
  preserved: [],
};

function skipped(call: string, title: string, why: string): LaneView {
  return {call, title, status: 'skipped', detail: why, elapsedMs: undefined, error: undefined};
}

function failed<T>(call: string, title: string, result: Lane<T>): LaneView {
  return {
    call,
    title,
    status: 'failed',
    detail: 'The library refused this input. Its own message is below.',
    elapsedMs: undefined,
    error: result.ok ? undefined : result.error,
  };
}

/**
 * A run, from whichever end the input arrived at.
 *
 * A sample enters at the write lane, because there is a workbook and no bytes yet. A dropped
 * file enters at the read lane, because there are bytes and no workbook. From there both go
 * through the same code, which is what makes the timings on the page comparable.
 */
function runFrom(
  input: {workbook: Workbook; sampleId: string} | {bytes: Uint8Array},
  label: string,
): RunState {
  const lanes: LaneView[] = [];
  let bytes: Uint8Array;

  if ('workbook' in input) {
    const written = write(input.workbook);
    if (!written.ok) {
      return {
        ...EMPTY,
        label,
        sampleId: input.sampleId,
        lanes: [
          failed('writeXlsx', 'Serialise the workbook', written),
          skipped('readXlsx', 'Read the package back', 'there are no bytes to read'),
          skipped('round trip', 'Read, write, read again', 'the first write failed'),
        ],
      };
    }
    bytes = written.value.bytes;
    lanes.push({
      call: 'writeXlsx',
      title: 'Serialise the workbook',
      status: 'ok',
      detail: `${formatBytes(written.value.byteLength)} package`,
      elapsedMs: written.value.elapsedMs,
      error: undefined,
    });
  } else {
    bytes = input.bytes;
    lanes.push(
      skipped(
        'writeXlsx',
        'Serialise the workbook',
        `your file arrived as ${formatBytes(bytes.length)}`,
      ),
    );
  }

  const readBack = read(bytes);
  if (!readBack.ok) {
    return {
      ...EMPTY,
      label,
      sampleId: 'workbook' in input ? input.sampleId : undefined,
      lanes: [
        ...lanes,
        failed('readXlsx', 'Read the package back', readBack),
        skipped('round trip', 'Read, write, read again', 'the read failed'),
      ],
    };
  }
  const cells = readBack.value.sheets.reduce((total, sheet) => total + sheet.rowCount, 0);
  lanes.push({
    call: 'readXlsx',
    title: 'Read the package back',
    status: 'ok',
    detail:
      `${formatCount(readBack.value.sheets.length)} sheet(s), ${formatCount(cells)} rows` +
      (readBack.value.preservedParts.length === 0
        ? ''
        : `, ${formatCount(readBack.value.preservedParts.length)} preserved part(s)`),
    elapsedMs: readBack.value.elapsedMs,
    error: undefined,
  });

  const trip = roundTrip(bytes);
  if (!trip.ok) {
    lanes.push(failed('round trip', 'Read, write, read again', trip));
  } else {
    lanes.push({
      call: 'round trip',
      title: 'Read, write, read again',
      status: trip.value.agrees ? 'ok' : 'failed',
      detail: trip.value.agrees
        ? `Both reads agree, and the two writes are byte for byte identical: ${formatBytes(trip.value.firstByteLength)}.`
        : `The two reads disagree at ${trip.value.firstDisagreement ?? 'an unnamed path'}.`,
      elapsedMs: trip.value.elapsedMs,
      error: undefined,
    });
  }

  const workbook = readBack.value.workbook;
  const firstSheet = workbook.worksheets[0];
  return {
    label,
    sampleId: 'workbook' in input ? input.sampleId : undefined,
    lanes,
    bytes,
    grid: firstSheet === undefined ? undefined : toGrid(firstSheet, DEFAULT_BOUNDS),
    sheetNames: workbook.worksheets.map((sheet) => sheet.name),
    parts: listParts(bytes),
    preserved: readBack.value.preservedParts,
  };
}

export default defineComponent({
  name: 'Playground',
  setup() {
    const state = shallowRef<RunState>(EMPTY);
    const panel = ref<Panel>('sheet');
    const sheetName = ref<string | undefined>(undefined);
    const dragging = ref(false);
    const failure = ref<string | undefined>(undefined);
    const downloadUrl = ref<string | undefined>(undefined);

    const revoke = (): void => {
      if (downloadUrl.value !== undefined) URL.revokeObjectURL(downloadUrl.value);
      downloadUrl.value = undefined;
    };

    const settle = (next: RunState): void => {
      revoke();
      state.value = next;
      sheetName.value = next.sheetNames[0];
      if (next.bytes !== undefined) {
        downloadUrl.value = URL.createObjectURL(
          new Blob([next.bytes as BlobPart], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          }),
        );
      }
    };

    const runSample = (id: string): void => {
      const sample = findSample(id);
      if (sample === undefined) return;
      failure.value = undefined;
      settle(runFrom({workbook: sample.build(), sampleId: id}, sample.title));
    };

    const runFile = (file: File): void => {
      failure.value = undefined;
      void file
        .arrayBuffer()
        .then((buffer) => {
          settle(runFrom({bytes: new Uint8Array(buffer)}, file.name));
        })
        .catch((err: unknown) => {
          // Reading the File itself failed, which is the browser's problem and not the
          // library's. Saying so keeps the lane boxes honest about whose error is whose.
          failure.value = `The browser could not read that file: ${err instanceof Error ? err.message : String(err)}`;
        });
    };

    const showSheet = (name: string): void => {
      sheetName.value = name;
      const current = state.value;
      if (current.bytes === undefined) return;
      const readBack = read(current.bytes);
      if (!readBack.ok) return;
      const sheet = readBack.value.workbook.getWorksheet(name);
      if (sheet !== undefined) {
        state.value = {...current, grid: toGrid(sheet, DEFAULT_BOUNDS)};
      }
    };

    const sourceHtml = computed(() =>
      state.value.sampleId === undefined ? undefined : SAMPLE_SOURCES[state.value.sampleId],
    );

    onBeforeUnmount(revoke);
    runSample(SAMPLES[0]?.id ?? '');

    const picker = (): ReturnType<typeof h> =>
      h('div', {class: 'pg-picker'}, [
        h('div', {class: 'xlsx-kicker'}, 'Pick a workbook'),
        h(
          'div',
          {class: 'pg-picker__row'},
          SAMPLES.map((sample) =>
            h(
              'button',
              {
                type: 'button',
                key: sample.id,
                class: ['pg-chip', sample.id === state.value.sampleId ? 'is-on' : undefined],
                onClick: () => {
                  runSample(sample.id);
                },
              },
              sample.title,
            ),
          ),
        ),
        h(
          'p',
          {class: 'pg-picker__blurb'},
          state.value.sampleId === undefined
            ? `Your file: ${state.value.label}`
            : (findSample(state.value.sampleId)?.description ?? ''),
        ),
      ]);

    const dropZone = (): ReturnType<typeof h> =>
      h(
        'label',
        {
          class: ['pg-drop', dragging.value ? 'is-over' : undefined],
          onDragover: (event: DragEvent) => {
            event.preventDefault();
            dragging.value = true;
          },
          onDragleave: () => {
            dragging.value = false;
          },
          onDrop: (event: DragEvent) => {
            event.preventDefault();
            dragging.value = false;
            const file = event.dataTransfer?.files[0];
            if (file !== undefined) runFile(file);
          },
        },
        [
          h('input', {
            type: 'file',
            accept: '.xlsx,.xlsm,.xlsb',
            class: 'pg-drop__input',
            onChange: (event: Event) => {
              const file = (event.target as HTMLInputElement).files?.[0];
              if (file !== undefined) runFile(file);
            },
          }),
          h(
            'span',
            'Drop a spreadsheet here, or choose one. It is read in this tab and nowhere else.',
          ),
        ],
      );

    const tabs = (): ReturnType<typeof h> =>
      h(
        'div',
        {class: 'pg-tabs'},
        (
          [
            ['sheet', 'The sheet'],
            ['parts', `Inside the package (${formatCount(state.value.parts.length)})`],
            ['source', 'The code that made it'],
          ] as const
        ).map(([id, label]) =>
          h(
            'button',
            {
              type: 'button',
              key: id,
              class: ['pg-tab', panel.value === id ? 'is-on' : undefined],
              onClick: () => {
                panel.value = id;
              },
            },
            label,
          ),
        ),
      );

    const sheetTabs = (): ReturnType<typeof h> | null => {
      if (state.value.sheetNames.length < 2) return null;
      return h(
        'div',
        {class: 'pg-sheettabs'},
        state.value.sheetNames.map((name) =>
          h(
            'button',
            {
              type: 'button',
              key: name,
              class: ['pg-chip', name === sheetName.value ? 'is-on' : undefined],
              onClick: () => {
                showSheet(name);
              },
            },
            name,
          ),
        ),
      );
    };

    const body = (): ReturnType<typeof h> => {
      const current = state.value;
      if (panel.value === 'parts') {
        return h(PartTree, {parts: current.parts, preserved: current.preserved});
      }
      if (panel.value === 'source') {
        return h(SourceBlock, {
          html: sourceHtml.value,
          caption: 'Read straight out of www/playground/samples.ts at build time.',
        });
      }
      if (current.grid === undefined) {
        return h('p', {class: 'pg-empty'}, 'Nothing was read, so there is no sheet to paint.');
      }
      return h('div', [sheetTabs(), h(SheetGrid, {grid: current.grid})]);
    };

    return () =>
      h('div', {class: 'pg'}, [
        h('div', {class: 'pg-controls'}, [picker(), dropZone()]),
        failure.value === undefined ? null : h('p', {class: 'pg-failure'}, failure.value),
        h(LaneStrip, {lanes: state.value.lanes}),
        downloadUrl.value === undefined
          ? null
          : h(
              'a',
              {
                class: 'pg-download',
                href: downloadUrl.value,
                download: `${state.value.sampleId ?? 'workbook'}.xlsx`,
              },
              'Download these bytes',
            ),
        tabs(),
        h('div', {class: 'pg-panel'}, body()),
      ]);
  },
});
