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
  assert.match(xml, /<fills count="2">/); // just the two reserved fills
  assert.match(xml, /<cellXfs count="1">/); // just the default xf
});
