// The write-side style table: the interned `<fills>`/`<cellXfs>` backing styles.xml.
//
// OOXML styles are a *shared* table referenced by index: a cell (or a formatted row)
// names a `<cellXfs>` entry via its `s` attribute, and that entry names a fill by id.
// Identical styles must collapse to one entry — both to produce well-formed OOXML and
// to keep write cost bounded on large, lightly-formatted sheets (the historical
// performance cliff came from re-serialising a distinct style per cell). The registry
// interns each distinct fill and each distinct xf, handing back a stable index.
//
// Only fills are modelled today; fonts, borders, number formats, alignment, and
// protection extend the xf signature as those slices land. Until then every xf shares
// the single default font/border, and an unstyled cell/row resolves to xf 0.

import type {Color, Fill} from '../../core/style.ts';
import {XML_DECLARATION} from './xml.ts';

// Excel reserves fill ids 0 and 1 for the "none" and "gray125" patterns it always emits;
// custom fills are numbered from 2 so a foreign reader's built-in assumptions still hold.
const RESERVED_FILL_COUNT = 2;

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DEFAULT_FONT =
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

// One interned cell format. Today it is fully described by its fill id (0 = no fill);
// added facets (fontId, borderId, numFmtId, …) become further fields and signature terms.
interface CellFormat {
  readonly fillId: number;
}

export class StyleRegistry {
  // Custom fill xml fragments, in id order; the emitted id is RESERVED_FILL_COUNT + index.
  readonly #fillXml: string[] = [];
  readonly #fillIdBySignature = new Map<string, number>();

  // xf 0 is the default (no fill); further entries are appended as distinct styles appear.
  readonly #formats: CellFormat[] = [{fillId: 0}];
  readonly #xfIndexBySignature = new Map<string, number>();

  /**
   * The `<cellXfs>` index for a cell or row carrying `fill`. A `none`/absent fill needs
   * no style and resolves to the default xf 0, so it emits no `s` attribute at all.
   */
  styleId(fill: Fill | undefined): number {
    if (fill === undefined || fill.pattern === 'none') return 0;
    const fillId = this.#internFill(fill);
    const signature = `fill:${fillId}`;
    let index = this.#xfIndexBySignature.get(signature);
    if (index === undefined) {
      index = this.#formats.length;
      this.#formats.push({fillId});
      this.#xfIndexBySignature.set(signature, index);
    }
    return index;
  }

  #internFill(fill: Fill): number {
    const signature = fillSignature(fill);
    let id = this.#fillIdBySignature.get(signature);
    if (id === undefined) {
      id = RESERVED_FILL_COUNT + this.#fillXml.length;
      this.#fillXml.push(fillXml(fill));
      this.#fillIdBySignature.set(signature, id);
    }
    return id;
  }

  /** Serialise the accumulated table into a complete, valid styles.xml part. */
  toXml(): string {
    const fillCount = RESERVED_FILL_COUNT + this.#fillXml.length;
    const fills =
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      this.#fillXml.join('');
    const cellXfs = this.#formats.map(cellXf).join('');
    return (
      XML_DECLARATION +
      `<styleSheet xmlns="${NS_MAIN}">` +
      `<fonts count="1">${DEFAULT_FONT}</fonts>` +
      `<fills count="${fillCount}">${fills}</fills>` +
      `<borders count="1">${DEFAULT_BORDER}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.#formats.length}">${cellXfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/>' +
      '</styleSheet>'
    );
  }
}

function cellXf(format: CellFormat): string {
  const applyFill = format.fillId !== 0 ? ' applyFill="1"' : '';
  return `<xf numFmtId="0" fontId="0" fillId="${format.fillId}" borderId="0" xfId="0"${applyFill}/>`;
}

// A stable, collision-free key for a fill: identical fills share it, distinct ones don't.
function fillSignature(fill: Fill): string {
  return `${fill.pattern}|${colorSignature(fill.fgColor)}|${colorSignature(fill.bgColor)}`;
}

function colorSignature(color: Color | undefined): string {
  if (color === undefined) return '';
  return `${color.argb ?? ''}/${color.theme ?? ''}/${color.tint ?? ''}/${color.indexed ?? ''}`;
}

function fillXml(fill: Fill): string {
  const fg = fill.fgColor ? `<fgColor ${colorAttrs(fill.fgColor)}/>` : '';
  // A solid fill's background is the automatic indexed placeholder unless one is stated;
  // omitting it makes Excel render the fill as flat black, so it is always emitted.
  const bg = fill.bgColor
    ? `<bgColor ${colorAttrs(fill.bgColor)}/>`
    : fill.pattern === 'solid'
      ? '<bgColor indexed="64"/>'
      : '';
  return `<fill><patternFill patternType="${fill.pattern}">${fg}${bg}</patternFill></fill>`;
}

function colorAttrs(color: Color): string {
  const parts: string[] = [];
  if (color.argb !== undefined) parts.push(`rgb="${color.argb}"`);
  if (color.theme !== undefined) parts.push(`theme="${color.theme}"`);
  if (color.tint !== undefined) parts.push(`tint="${color.tint}"`);
  if (color.indexed !== undefined) parts.push(`indexed="${color.indexed}"`);
  return parts.join(' ');
}
