import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {Fill} from '../../core/style.ts';
import {StyleRegistry} from './styles.ts';

const solid = (argb: string): Fill => ({type: 'pattern', pattern: 'solid', fgColor: {argb}});

test('an absent or "none" fill resolves to the default xf 0 — no style entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId(undefined), 0);
  assert.equal(styles.styleId({type: 'pattern', pattern: 'none'}), 0);
});

test('an identical fill interns to one shared xf index however many times it is seen', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId(solid('FFFF0000'));
  for (let i = 0; i < 40; i++) {
    assert.equal(styles.styleId(solid('FFFF0000')), first, 'every identical fill returns the same index');
  }
  assert.notEqual(first, 0, 'a real fill gets a non-default index');
});

test('genuinely different fills get distinct xf indices — dedup does not over-collapse', () => {
  const styles = new StyleRegistry();
  const red = styles.styleId(solid('FFFF0000'));
  const blue = styles.styleId(solid('FF0000FF'));
  const redPattern = styles.styleId({type: 'pattern', pattern: 'lightGrid', fgColor: {argb: 'FFFF0000'}});
  assert.equal(new Set([red, blue, redPattern]).size, 3, 'three visually distinct fills, three indices');
});

test('the emitted stylesheet reflects the interned fills and cell formats', () => {
  const styles = new StyleRegistry();
  styles.styleId(solid('FFFF0000'));
  styles.styleId(solid('FFFF0000')); // deduped
  styles.styleId(solid('FF00FF00'));
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

test('an empty registry still emits a valid minimal stylesheet', () => {
  const xml = new StyleRegistry().toXml();
  assert.match(xml, /<fills count="2">/); // just the two reserved fills
  assert.match(xml, /<cellXfs count="1">/); // just the default xf
});
