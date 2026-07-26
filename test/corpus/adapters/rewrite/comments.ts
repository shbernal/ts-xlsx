// Threaded comments, legacy notes, and rich text.

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';
import type {CorpusApi} from '../../case.ts';
import {commentThreadFacts, packagePartFacts, partMapOf} from './package-facts.ts';
import {readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';

export const comments = {
  // Write a noted cell, relocate its comments part to a non-canonical path (xl/sheet1_comments.xml)
  // reachable only through the worksheet rels, and reload → { ok, error, note }. The reader locates the
  // comments part by relationship *type*, not by filename glob, so the moved part still loads and its
  // note reads back.
  nonCanonicalCommentsPartReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.getCell('A1').note = 'hi';
    const files = unzipSync(new Uint8Array(writeXlsx(wb)));
    const commentPart = Object.keys(files).find((n) => /^xl\/comments\d*\.xml$/.test(n))!;
    const relName = 'xl/worksheets/_rels/sheet1.xml.rels';
    files['xl/sheet1_comments.xml'] = files[commentPart]!;
    delete files[commentPart];
    files[relName] = strToU8(
      strFromU8(files[relName]!).replace(
        /Target="[^"]*comments\d*\.xml"/i,
        'Target="../sheet1_comments.xml"',
      ),
    );
    const buffer = zipSync(files);
    let ok = false;
    let error = null;
    let note = null;
    try {
      note = readXlsx(buffer).worksheets[0]!.getCell('A1').note ?? null;
      ok = true;
    } catch (e) {
      error = String((e as CorpusApi)?.message || e);
    }
    return {ok, error, note};
  },

  // Attach then clear a cell note and report whether the written package still carries a comment/VML
  // artifact → { commentPartPresent, vmlPartPresent, readNoteAfter, neighborNoteIntact,
  // cleanHasCommentPart }. Clearing a note (note = undefined) is a genuine removal: the sole-noted
  // cell's comment part disappears and reads back null, while a neighbor that kept its note is intact,
  // and a workbook that never had a note emits no comment part at all.
  removeCellNoteReport() {
    const partNames = (wb: CorpusApi) => Object.keys(partMapOf(writeXlsx(wb)));
    // The note being cleared is the only one, so a lingering comment part would prove the removal
    // merely emptied the note rather than dropping it.
    const solo = new Workbook();
    const soloSheet = solo.addWorksheet('S');
    soloSheet.getCell('A1').value = 'x';
    soloSheet.getCell('A1').note = 'remove me';
    soloSheet.getCell('A1').note = undefined;
    const soloParts = partNames(solo);
    const commentPartPresent = soloParts.some((f) => /comments\d*\.xml$/.test(f));
    const vmlPartPresent = soloParts.some((f) => /vmlDrawing\d*\.vml$/.test(f));
    const readNoteAfter = readXlsx(writeXlsx(solo)).worksheets[0]!.getCell('A1').note ?? null;

    // Clearing one note must not disturb another cell's note.
    const pair = new Workbook();
    const pairSheet = pair.addWorksheet('S');
    pairSheet.getCell('A1').value = 'x';
    pairSheet.getCell('A1').note = 'remove me';
    pairSheet.getCell('B1').value = 'y';
    pairSheet.getCell('B1').note = 'keep me';
    pairSheet.getCell('A1').note = undefined;
    const reloaded = readXlsx(writeXlsx(pair)).worksheets[0]!;
    const neighborNoteIntact = !!reloaded.getCell('B1').note;

    const clean = new Workbook();
    clean.addWorksheet('S').getCell('A1').value = 'x';
    const cleanHasCommentPart = partNames(clean).some((f) => /comments\d*\.xml$/.test(f));

    return {
      commentPartPresent,
      vmlPartPresent,
      readNoteAfter,
      neighborNoteIntact,
      cleanHasCommentPart,
    };
  },

  // Write a rich-text cell, then read it back, and report how its runs serialized and survived →
  // { emptyTextRunInXml, runCount, runs: [{text, bold, italic, underline}] }. Mirrors the oracle. A
  // zero-length run must never emit an empty <t> (Excel flags it corrupt); the surviving runs keep
  // their text and per-run formatting. The rewrite writes rich text inline, so both the empty-<t>
  // scan and the read-back target the worksheet XML (there is no shared-strings part).
  async richTextRoundtripReport(runs: CorpusApi) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = {richText: runs};
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const xml = parts['xl/sharedStrings.xml'] || parts['xl/worksheets/sheet1.xml'] || '';
    const emptyTextRunInXml = /<(?:\w+:)?t\b[^>]*\/>|<(?:\w+:)?t\b[^>]*><\/(?:\w+:)?t>/.test(xml);
    const value = readXlsx(buffer).getWorksheet('S')!.getCell('A1').value as CorpusApi;
    const readRuns = value && Array.isArray(value.richText) ? value.richText : [];
    return {
      emptyTextRunInXml,
      runCount: readRuns.length,
      runs: readRuns.map((r: CorpusApi) => ({
        text: r.text ?? null,
        bold: r.font ? (r.font.bold ?? false) : false,
        italic: r.font ? (r.font.italic ?? false) : false,
        underline: r.font ? (r.font.underline ?? false) : false,
      })),
    };
  },

  // Read a fixture and report the modern threaded conversations the reader reconstructs — see
  // {@link commentThreadFacts} for the shape.
  readFixtureCommentThreads(rel: CorpusApi, refs: CorpusApi = []) {
    return commentThreadFacts(readFixture(rel), refs);
  },

  // Author a conversation in the MODEL — never read from a file — write it, and read the package back:
  // → { parts, deterministic, model }. Nothing here comes from preserved bytes, so this is what proves the
  // writer serialises the feature rather than merely carrying it: the thread part, the workbook person
  // registry, and the legacy fallback comment are all built from the model on the way out and reassembled
  // into threads on the way in. `deterministic` re-writes the same model and compares the bytes — the
  // writer holds no clock and no id generator, so every guid and timestamp comes from the caller.
  //
  // The conversation deliberately mixes every facet at once (a reply by a second author, an @mention
  // resolving through a third registry entry, a resolved thread, and a genuine legacy note on another
  // cell) because those are the facets that interact: the fallback text folds the reply in, the mention
  // spans the message text, and the note must not be mistaken for the thread's fallback.
  authoredCommentThreadRoundtrip() {
    const ADA = '{39236F6F-643D-4654-8264-DD21C8472F7F}';
    const GRACE = '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}';
    const GRACE_MENTIONED = '{BA397017-DD76-4496-AA75-59ADB199950C}';
    const workbook = new Workbook();
    workbook.addPerson({id: ADA, displayName: 'Ada Lovelace', providerId: 'AD'});
    workbook.addPerson({id: GRACE, displayName: 'Grace Hopper', providerId: 'AD'});
    // The same human as GRACE under a second id — how Excel registers a *mentioned* identity.
    workbook.addPerson({
      id: GRACE_MENTIONED,
      displayName: 'Grace Hopper',
      providerId: 'PeoplePicker',
    });
    const sheet = workbook.addWorksheet('Review');
    sheet.getCell('B2').value = 1234;
    sheet.getCell('D4').note = 'an ordinary note beside the conversation';
    sheet.addCommentThread({
      // Absolute on purpose: the anchor is canonicalised, so `$B$2` and `B2` are one cell.
      ref: '$B$2',
      resolved: true,
      comments: [
        {
          id: '{11111111-2222-3333-4444-555555555555}',
          personId: ADA,
          date: '2026-07-26T12:00:00.00',
          text: '@Grace Hopper is this gross or net?',
          mentions: [
            {
              personId: GRACE_MENTIONED,
              mentionId: '{3F2C1A9E-5B84-4D67-9C2E-71A0D4E8B531}',
              startIndex: 0,
              length: 13,
            },
          ],
        },
        {
          id: '{66666666-7777-8888-9999-AAAAAAAAAAAA}',
          personId: GRACE,
          date: '2026-07-26T12:05:00.00',
          text: 'Gross. Confirmed with finance.',
          mentions: [],
        },
      ],
    });
    const bytes = writeXlsx(workbook);
    return {
      parts: packagePartFacts(partMapOf(bytes)),
      deterministic: Buffer.compare(Buffer.from(bytes), Buffer.from(writeXlsx(workbook))) === 0,
      model: commentThreadFacts(readXlsx(bytes), ['B2', '$B$2', 'D4']),
    };
  },
};
