// The write-side style table: the interned `<numFmts>`/`<fills>`/`<cellXfs>` backing styles.xml.
//
// OOXML styles are a *shared* table referenced by index: a cell (or a formatted row/column)
// names a `<cellXfs>` entry via its `s` attribute, and that entry names a fill by id and a
// number format by id. Identical styles must collapse to one entry — both to produce
// well-formed OOXML and to keep write cost bounded on large, lightly-formatted sheets (the
// historical performance cliff came from re-serialising a distinct style per cell). The
// registry interns each distinct fill, number format, and xf, handing back a stable index.
//
// Only fills and number formats are modelled today; fonts, borders, alignment, and
// protection extend the xf signature as those slices land. Until then every xf shares the
// single default font/border, and an unstyled cell/row/column resolves to xf 0.

import type {Color, Fill, Font, UnderlineStyle} from '../../core/style.ts';
import {escapeAttr, XML_DECLARATION} from './xml.ts';

// Excel reserves fill ids 0 and 1 for the "none" and "gray125" patterns it always emits;
// custom fills are numbered from 2 so a foreign reader's built-in assumptions still hold.
const RESERVED_FILL_COUNT = 2;

// Font id 0 is the always-present default font (Calibri 11); custom fonts are numbered from 1.
const RESERVED_FONT_COUNT = 1;

// numFmt ids below 164 are reserved by ECMA-376 for the built-in formats every consumer
// knows implicitly; custom format codes are numbered from 164 up. Id 0 is General (no code).
const CUSTOM_NUMFMT_BASE = 164;

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DEFAULT_FONT =
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

/** A cell's style facets as the writer composes them: cell overrides atop row/column defaults. */
export interface CellStyle {
  readonly fill?: Fill | undefined;
  readonly numFmt?: string | undefined;
  readonly font?: Partial<Font> | undefined;
}

// One interned cell format. `fillId` 0 is no fill; `numFmtId` 0 is the General format;
// `fontId` 0 is the default font. Added facets (borderId, …) become further fields and terms.
interface CellFormat {
  readonly fillId: number;
  readonly numFmtId: number;
  readonly fontId: number;
}

export class StyleRegistry {
  // Custom fill xml fragments, in id order; the emitted id is RESERVED_FILL_COUNT + index.
  readonly #fillXml: string[] = [];
  readonly #fillIdBySignature = new Map<string, number>();

  // Custom number-format codes, in id order; the emitted id is CUSTOM_NUMFMT_BASE + index.
  readonly #numFmtCodes: string[] = [];
  readonly #numFmtIdByCode = new Map<string, number>();

  // Custom font xml fragments, in id order; the emitted id is RESERVED_FONT_COUNT + index.
  readonly #fontXml: string[] = [];
  readonly #fontIdBySignature = new Map<string, number>();

  // xf 0 is the default (no fill/font, General format); further entries append as styles appear.
  readonly #formats: CellFormat[] = [{fillId: 0, numFmtId: 0, fontId: 0}];
  readonly #xfIndexBySignature = new Map<string, number>();

  /**
   * The `<cellXfs>` index for a composed cell/row/column style. A style with no facet needs
   * no entry and resolves to the default xf 0, so its owner emits no `s` attribute at all.
   */
  styleId(style: CellStyle): number {
    const fillId = style.fill && style.fill.pattern !== 'none' ? this.#internFill(style.fill) : 0;
    const numFmtId = style.numFmt !== undefined && style.numFmt !== '' ? this.#internNumFmt(style.numFmt) : 0;
    const fontId = style.font ? this.#internFont(style.font) : 0;
    if (fillId === 0 && numFmtId === 0 && fontId === 0) return 0;

    const signature = `fill:${fillId}|numFmt:${numFmtId}|font:${fontId}`;
    let index = this.#xfIndexBySignature.get(signature);
    if (index === undefined) {
      index = this.#formats.length;
      this.#formats.push({fillId, numFmtId, fontId});
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

  #internNumFmt(code: string): number {
    let id = this.#numFmtIdByCode.get(code);
    if (id === undefined) {
      id = CUSTOM_NUMFMT_BASE + this.#numFmtCodes.length;
      this.#numFmtCodes.push(code);
      this.#numFmtIdByCode.set(code, id);
    }
    return id;
  }

  // A font whose partial carries no facet that differs from the default contributes nothing
  // and maps to font id 0; otherwise its serialised form is interned and dedup'd like a fill.
  #internFont(font: Partial<Font>): number {
    const xml = fontXml(font);
    if (xml === '') return 0;
    let id = this.#fontIdBySignature.get(xml);
    if (id === undefined) {
      id = RESERVED_FONT_COUNT + this.#fontXml.length;
      this.#fontXml.push(`<font>${xml}</font>`);
      this.#fontIdBySignature.set(xml, id);
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
    const fontCount = RESERVED_FONT_COUNT + this.#fontXml.length;
    const fonts = DEFAULT_FONT + this.#fontXml.join('');
    return (
      XML_DECLARATION +
      `<styleSheet xmlns="${NS_MAIN}">` +
      this.#numFmtsXml() +
      `<fonts count="${fontCount}">${fonts}</fonts>` +
      `<fills count="${fillCount}">${fills}</fills>` +
      `<borders count="1">${DEFAULT_BORDER}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.#formats.length}">${cellXfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/>' +
      '</styleSheet>'
    );
  }

  // <numFmts> is the first child of <styleSheet> and is omitted entirely when no custom
  // format was used, matching how Excel writes an all-built-in workbook.
  #numFmtsXml(): string {
    if (this.#numFmtCodes.length === 0) return '';
    const entries = this.#numFmtCodes
      .map((code, i) => `<numFmt numFmtId="${CUSTOM_NUMFMT_BASE + i}" formatCode="${escapeFormatCode(code)}"/>`)
      .join('');
    return `<numFmts count="${this.#numFmtCodes.length}">${entries}</numFmts>`;
  }
}

function cellXf(format: CellFormat): string {
  const applyFill = format.fillId !== 0 ? ' applyFill="1"' : '';
  const applyNumberFormat = format.numFmtId !== 0 ? ' applyNumberFormat="1"' : '';
  const applyFont = format.fontId !== 0 ? ' applyFont="1"' : '';
  return (
    `<xf numFmtId="${format.numFmtId}" fontId="${format.fontId}" fillId="${format.fillId}" borderId="0" xfId="0"` +
    `${applyNumberFormat}${applyFont}${applyFill}/>`
  );
}

// Serialise the facets a cell font overrides, in ECMA-376 CT_Font child order. A boolean
// flag is emitted only when true (its absence is the default false); an empty result means
// the font differs from the default in nothing and needs no <fonts> entry at all.
function fontXml(font: Partial<Font>): string {
  const parts: string[] = [];
  if (font.bold) parts.push('<b/>');
  if (font.italic) parts.push('<i/>');
  if (font.strike) parts.push('<strike/>');
  if (font.outline) parts.push('<outline/>');
  const underline = underlineXml(font.underline);
  if (underline !== '') parts.push(underline);
  if (font.vertAlign !== undefined) parts.push(`<vertAlign val="${font.vertAlign}"/>`);
  if (font.size !== undefined) parts.push(`<sz val="${numberAttr(font.size)}"/>`);
  if (font.color !== undefined) parts.push(`<color ${colorAttrs(font.color)}/>`);
  if (font.name !== undefined) parts.push(`<name val="${escapeAttr(font.name)}"/>`);
  if (font.family !== undefined) parts.push(`<family val="${numberAttr(font.family)}"/>`);
  if (font.charset !== undefined) parts.push(`<charset val="${numberAttr(font.charset)}"/>`);
  if (font.scheme !== undefined && font.scheme !== 'none') parts.push(`<scheme val="${font.scheme}"/>`);
  return parts.join('');
}

// `<u/>` is single underline (the same as an explicit "single"); the named variants carry a
// val; false and "none" are the default no-underline and emit nothing.
function underlineXml(underline: UnderlineStyle | undefined): string {
  if (underline === undefined || underline === false || underline === 'none') return '';
  if (underline === true || underline === 'single') return '<u/>';
  return `<u val="${underline}"/>`;
}

function numberAttr(value: number): string {
  if (!Number.isFinite(value)) {
    throw new Error(`cannot serialise a non-finite font metric (${value})`);
  }
  return String(value);
}

// A format code sits in the `formatCode` attribute; only the markup-significant characters
// need escaping. A code can legitimately contain `"` (quoted literals like `"$"`), `<`, `&`.
function escapeFormatCode(code: string): string {
  return code
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
