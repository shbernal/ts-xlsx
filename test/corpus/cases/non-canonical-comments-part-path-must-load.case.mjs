export default {
  id: 'non-canonical-comments-part-path-must-load',
  cluster: 'address-decoding',
  description:
    'Open Packaging Conventions let a part live at whatever path its relationship Target names — the ' +
    'comments part need not be the conventional xl/comments1.xml. A reader must locate parts by ' +
    'relationship type, not by filename glob. When the comments part sits at a non-canonical path ' +
    '(xl/sheet1_comments.xml) reachable only through the worksheet rels, a glob-based loader skips ' +
    'it and then crashes reconciling the rels; the package must load without throwing.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'a workbook whose comments part lives at a non-canonical path loads without throwing',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, error} = await api.nonCanonicalCommentsPartReport();
        assert.strictEqual(error, null, `load must not throw (got: ${error})`);
        assert.strictEqual(ok, true);
      },
    },
    {
      name: 'the comment carried by the non-canonically-located part is still read',
      baseline: 'fail',
      async expect(api, assert) {
        const {note} = await api.nonCanonicalCommentsPartReport();
        assert.strictEqual(note && note.texts ? note.texts.map(t => t.text).join('') : note, 'hi');
      },
    },
  ],
};
