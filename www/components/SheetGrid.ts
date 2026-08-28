/**
 * A worksheet, painted.
 *
 * The model decided what to show; this decides only how it looks. What it paints is the
 * value's text with the cell's own font, fill and alignment applied, and the number format
 * shown beside the cell rather than applied to it: the library models formats and does not
 * implement a formatting engine, so a grid that rendered `$1,234.50` here would be showing
 * a cell this library cannot produce.
 */

import {defineComponent, h, type PropType} from 'vue';

import type {Alignment, Fill, Font} from '../../src/index.ts';
import {formatCount} from '../playground/format.ts';
import type {DisplayCell, Grid} from '../playground/grid.ts';

/** An ARGB from the model, as a CSS colour. The alpha byte leads, so it moves to the end. */
function cssColor(argb: string | undefined): string | undefined {
  if (argb === undefined || !/^[0-9A-Fa-f]{8}$/.test(argb)) return undefined;
  return `#${argb.slice(2)}${argb.slice(0, 2)}`;
}

function fontStyle(font: Font | undefined): Record<string, string> {
  if (font === undefined) return {};
  const style: Record<string, string> = {};
  if (font.bold === true) style['fontWeight'] = '600';
  if (font.italic === true) style['fontStyle'] = 'italic';
  if (font.strike === true) style['textDecoration'] = 'line-through';
  if (font.size !== undefined) style['fontSize'] = `${font.size / 11}em`;
  const color = cssColor(font.color?.argb);
  if (color !== undefined) style['color'] = color;
  return style;
}

function fillStyle(fill: Fill | undefined): Record<string, string> {
  // Only a solid pattern is painted. The hatch patterns are carried for fidelity on read,
  // and drawing an approximation of one would be inventing a rendering nobody wrote.
  if (fill === undefined || fill.type !== 'pattern' || fill.pattern !== 'solid') return {};
  const color = cssColor(fill.fgColor?.argb);
  return color === undefined ? {} : {backgroundColor: color};
}

function alignmentStyle(
  alignment: Alignment | undefined,
  numeric: boolean,
): Record<string, string> {
  const style: Record<string, string> = {};
  // Excel's own default: a number sits right, everything else left. Applied only where the
  // cell states nothing, so an authored alignment always wins.
  style['textAlign'] = alignment?.horizontal ?? (numeric ? 'right' : 'left');
  if (alignment?.vertical !== undefined) {
    style['verticalAlign'] = alignment.vertical === 'center' ? 'middle' : alignment.vertical;
  }
  if (alignment?.wrapText === true) style['whiteSpace'] = 'pre-wrap';
  return style;
}

function cellNode(cell: DisplayCell): ReturnType<typeof h> {
  const numeric = cell.type === 'number' || cell.type === 'date';
  return h(
    'td',
    {
      class: ['pg-cell', `pg-cell--${cell.type}`],
      // The number format is a fact about the cell, so it is offered on hover rather than
      // applied to the text. `title` is the one place it can go without inventing a column.
      title: cell.numFmt === undefined ? undefined : `${cell.address}  ${cell.numFmt}`,
      colspan: cell.merge === undefined ? undefined : cell.merge.colSpan,
      rowspan: cell.merge === undefined ? undefined : cell.merge.rowSpan,
      style: {
        ...fontStyle(cell.font),
        ...fillStyle(cell.fill),
        ...alignmentStyle(cell.alignment, numeric),
      },
    },
    cell.text,
  );
}

export default defineComponent({
  name: 'SheetGrid',
  props: {
    grid: {type: Object as PropType<Grid>, required: true},
  },
  setup(props) {
    return () => {
      const {grid} = props;
      if (grid.rows.length === 0) {
        return h('p', {class: 'pg-empty'}, `${grid.sheetName} holds no cells.`);
      }
      const clipped =
        grid.hiddenRows === 0 && grid.hiddenColumns === 0
          ? null
          : h(
              'p',
              {class: 'pg-clipped'},
              `Showing ${formatCount(grid.rows.length)} of ${formatCount(grid.totalRows)} rows and ` +
                `${formatCount(grid.columns.length)} of ${formatCount(grid.totalColumns)} columns. ` +
                'The rest was read; it is not painted.',
            );
      return h('div', {class: 'pg-grid'}, [
        h('div', {class: 'pg-grid__scroll'}, [
          h('table', {class: 'pg-sheet'}, [
            h('thead', [
              h('tr', [
                h('th', {class: 'pg-corner'}, ''),
                ...grid.columns.map((letter) =>
                  h('th', {class: 'pg-colhead', key: letter}, letter),
                ),
              ]),
            ]),
            h(
              'tbody',
              grid.rows.map((row) =>
                h('tr', {key: row.number}, [
                  h('th', {class: 'pg-rowhead xlsx-figure'}, String(row.number)),
                  ...row.cells.map((cell) => cellNode(cell)),
                ]),
              ),
            ),
          ]),
        ]),
        clipped,
      ]);
    };
  },
});
