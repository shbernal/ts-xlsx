import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {Fill} from '../../core/style.ts';
import {StyleRegistry} from './styles.ts';

const solid = (argb: string): Fill => ({type: 'pattern', pattern: 'solid', fgColor: {argb}});

test('an absent or "none" fill with no format resolves to the default xf 0 — no style entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({}), 0);
  assert.equal(styles.styleId({fill: {type: 'pattern', pattern: 'none'}}), 0);
  assert.equal(styles.styleId({numFmt: ''}), 0);
});

test('an identical fill interns to one shared xf index however many times it is seen', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({fill: solid('FFFF0000')});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({fill: solid('FFFF0000')}), first, 'every identical fill returns the same index');
  }
  assert.notEqual(first, 0, 'a real fill gets a non-default index');
});

test('an identical number format interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({numFmt: '0.00%'});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({numFmt: '0.00%'}), first, 'every identical numFmt returns the same index');
  }
  assert.notEqual(first, 0, 'a real numFmt gets a non-default index');
});

test('fill and number format are independent facets — the same fill under two formats is two xfs', () => {
  const styles = new StyleRegistry();
  const plain = styles.styleId({fill: solid('FFFF0000')});
  const formatted = styles.styleId({fill: solid('FFFF0000'), numFmt: '0.00'});
  assert.notEqual(plain, formatted, 'adding a number format to a fill is a distinct style');
});

test('genuinely different styles get distinct xf indices — dedup does not over-collapse', () => {
  const styles = new StyleRegistry();
  const red = styles.styleId({fill: solid('FFFF0000')});
  const blue = styles.styleId({fill: solid('FF0000FF')});
  const pct = styles.styleId({numFmt: '0.00%'});
  const cur = styles.styleId({numFmt: '"$"#,##0.00'});
  assert.equal(new Set([red, blue, pct, cur]).size, 4, 'four distinct styles, four indices');
});

test('the emitted stylesheet reflects the interned fills and cell formats', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('FFFF0000')});
  styles.styleId({fill: solid('FFFF0000')}); // deduped
  styles.styleId({fill: solid('FF00FF00')});
  const xml = styles.toXml();

  // Two reserved fills (none, gray125) + two custom = 4; the default xf + two custom = 3.
  assert.match(xml, /<fills count="4">/);
  assert.match(xml, /<cellXfs count="3">/);
  // The visible colour is the pattern foreground with an automatic indexed background.
  assert.match(xml, /<patternFill patternType="solid"><fgColor rgb="FFFF0000"\/><bgColor indexed="64"\/><\/patternFill>/);
  // A styled xf references its fill and flags applyFill; the default xf does neither.
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"\/>/);
});

test('a custom number format is defined in <numFmts> from id 164 and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({numFmt: '0.00%'});
  styles.styleId({numFmt: '"$"#,##0.00'});
  const xml = styles.toXml();

  assert.match(xml, /<numFmts count="2">/);
  assert.match(xml, /<numFmt numFmtId="164" formatCode="0.00%"\/>/);
  // A quoted currency literal survives with its markup-significant characters escaped.
  assert.match(xml, /<numFmt numFmtId="165" formatCode="&quot;\$&quot;#,##0.00"\/>/);
  // The referencing xf names the custom id and flags applyNumberFormat.
  assert.match(xml, /<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"\/>/);
});

test('an empty registry still emits a valid minimal stylesheet with no <numFmts>', () => {
  const xml = new StyleRegistry().toXml();
  assert.doesNotMatch(xml, /<numFmts/); // omitted entirely when all-built-in
  assert.match(xml, /<fonts count="1">/); // just the default font
  assert.match(xml, /<fills count="2">/); // just the two reserved fills
  assert.match(xml, /<borders count="1">/); // just the empty default border
  assert.match(xml, /<cellXfs count="1">/); // just the default xf
});

test('an identical font interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({font: {bold: true}});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({font: {bold: true}}), first, 'every identical font returns the same index');
  }
  assert.notEqual(first, 0, 'a real font gets a non-default index');
});

test('a font that overrides nothing resolves to the default xf 0 — no font entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({font: {}}), 0);
  // A boolean flag that is explicitly false is the default and adds no <font>.
  assert.equal(styles.styleId({font: {bold: false, italic: false}}), 0);
  assert.doesNotMatch(styles.toXml(), /<b\/>/);
  assert.match(styles.toXml(), /<fonts count="1">/);
});

test('font, fill, and number format are independent facets composed into one xf', () => {
  const styles = new StyleRegistry();
  const bold = styles.styleId({font: {bold: true}});
  const boldRed = styles.styleId({font: {bold: true}, fill: solid('FFFF0000')});
  const boldRedPct = styles.styleId({font: {bold: true}, fill: solid('FFFF0000'), numFmt: '0.00%'});
  assert.equal(new Set([bold, boldRed, boldRedPct]).size, 3, 'each added facet is a distinct composed style');
});

test('a custom font is defined in <fonts> after the default and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({font: {bold: true, italic: true, size: 14, color: {argb: 'FF3A80D5'}, name: 'Arial'}});
  const xml = styles.toXml();

  // Default font (id 0) + one custom (id 1).
  assert.match(xml, /<fonts count="2">/);
  // The facets serialise in ECMA-376 child order, with the typeface attribute-escaped.
  assert.match(
    xml,
    /<font><b\/><i\/><sz val="14"\/><color rgb="FF3A80D5"\/><name val="Arial"\/><\/font>/
  );
  // The referencing xf names font id 1 and flags applyFont; the default xf keeps font id 0.
  assert.match(xml, /<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"\/>/);
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('an underline flag serialises single as a bare tag and a named variant with its val', () => {
  const styles = new StyleRegistry();
  styles.styleId({font: {underline: true}});
  styles.styleId({font: {underline: 'double'}});
  const xml = styles.toXml();
  assert.match(xml, /<font><u\/><\/font>/); // true → single → bare <u/>
  assert.match(xml, /<font><u val="double"\/><\/font>/);
});

test('an identical border interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({border: {top: {style: 'thin'}}});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({border: {top: {style: 'thin'}}}), first, 'every identical border returns the same index');
  }
  assert.notEqual(first, 0, 'a real border gets a non-default index');
});

test('a border that styles no edge resolves to the default xf 0 — no border entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({border: {}}), 0);
  assert.match(styles.toXml(), /<borders count="1">/); // just the empty default border
});

test('font, fill, number format, and border are independent facets composed into one xf', () => {
  const styles = new StyleRegistry();
  const b = styles.styleId({border: {top: {style: 'thin'}}});
  const bFill = styles.styleId({border: {top: {style: 'thin'}}, fill: solid('FFFF0000')});
  const bFillFont = styles.styleId({border: {top: {style: 'thin'}}, fill: solid('FFFF0000'), font: {bold: true}});
  assert.equal(new Set([b, bFill, bFillFont]).size, 3, 'each added facet is a distinct composed style');
});

test('a custom border is defined in <borders> after the default and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({border: {top: {style: 'thin'}, bottom: {style: 'medium', color: {argb: 'FF3A80D5'}}}});
  const xml = styles.toXml();

  // Empty default border (id 0) + one custom (id 1).
  assert.match(xml, /<borders count="2">/);
  // Edges serialise in schema order (left, right, top, bottom, diagonal); a styleless edge is
  // a bare self-closing tag, a styled one carries its style and any colour child.
  assert.match(
    xml,
    /<border><left\/><right\/><top style="thin"\/><bottom style="medium"><color rgb="FF3A80D5"\/><\/bottom><diagonal\/><\/border>/
  );
  // The referencing xf names border id 1 and flags applyBorder; the default xf keeps border id 0.
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"\/>/);
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('an identical alignment interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({alignment: {horizontal: 'center'}});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({alignment: {horizontal: 'center'}}), first, 'every identical alignment returns the same index');
  }
  assert.notEqual(first, 0, 'a real alignment gets a non-default index');
});

test('an all-default alignment resolves to the default xf 0 — no <alignment>', () => {
  const styles = new StyleRegistry();
  // `general` horizontal is the default, and boolean flags left off contribute nothing.
  assert.equal(styles.styleId({alignment: {horizontal: 'general', wrapText: false, shrinkToFit: false}}), 0);
  assert.doesNotMatch(styles.toXml(), /<alignment/);
});

test('alignment composes into the xf as a child element, not a shared sub-table', () => {
  const styles = new StyleRegistry();
  styles.styleId({alignment: {horizontal: 'center', vertical: 'top', wrapText: true, indent: 2, textRotation: 45}});
  const xml = styles.toXml();

  // The aligned xf carries an <alignment> child in ECMA-376 attribute order and flags applyAlignment;
  // it is no longer self-closing. The default xf stays self-closing with no alignment.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="top" textRotation="45" wrapText="1" indent="2"\/><\/xf>/
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('alignment is an independent facet composed alongside fill, font, and border', () => {
  const styles = new StyleRegistry();
  const a = styles.styleId({alignment: {wrapText: true}});
  const aFill = styles.styleId({alignment: {wrapText: true}, fill: solid('FFFF0000')});
  const aFillFont = styles.styleId({alignment: {wrapText: true}, fill: solid('FFFF0000'), font: {bold: true}});
  assert.equal(new Set([a, aFill, aFillFont]).size, 3, 'each added facet is a distinct composed style');
});

test('an identical protection interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({protection: {locked: false}});
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId({protection: {locked: false}}), first, 'every identical protection returns the same index');
  }
  assert.notEqual(first, 0, 'a real protection gets a non-default index');
});

test('an all-default protection resolves to the default xf 0 — no <protection>', () => {
  const styles = new StyleRegistry();
  // locked defaults to TRUE and hidden to false in OOXML, so a locked, non-hidden cell restates
  // the default and carries no information — it must not spend an xf entry.
  assert.equal(styles.styleId({protection: {locked: true, hidden: false}}), 0);
  assert.doesNotMatch(styles.toXml(), /<protection/);
});

test('protection composes into the xf as a child element, flagging applyProtection', () => {
  const styles = new StyleRegistry();
  styles.styleId({protection: {locked: false, hidden: true}});
  const xml = styles.toXml();

  // Only the meaningful flags serialise: the unlocked cell writes locked="0", the hidden one hidden="1".
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"\/><\/xf>/
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('alignment and protection compose as two xf children in schema order', () => {
  const styles = new StyleRegistry();
  styles.styleId({alignment: {horizontal: 'center'}, protection: {locked: false}});
  const xml = styles.toXml();

  // <alignment> precedes <protection> in the xf body, and both apply flags are set.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="center"\/><protection locked="0"\/><\/xf>/
  );
});

test('protection is an independent facet composed alongside fill, font, and border', () => {
  const styles = new StyleRegistry();
  const p = styles.styleId({protection: {locked: false}});
  const pFill = styles.styleId({protection: {locked: false}, fill: solid('FFFF0000')});
  const pFillFont = styles.styleId({protection: {locked: false}, fill: solid('FFFF0000'), font: {bold: true}});
  assert.equal(new Set([p, pFill, pFillFont]).size, 3, 'each added facet is a distinct composed style');
});
