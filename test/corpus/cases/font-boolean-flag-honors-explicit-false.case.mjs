// Cluster: styles
//
// Real-world scenario: OOXML boolean font flags (bold, italic, strike, condensed, outline, shadow,
// extend) default to ON when the tag is present with no value — `<b/>` means bold. But a tag can
// carry an explicit value: `<b val="0"/>` means bold is OFF, `<b val="1"/>` means ON. A parser that
// treats mere tag presence as true — ignoring the val attribute — reads an explicit-false flag as
// true, corrupting styles. This bites hardest in conditional-formatting differential (dxf) fonts,
// where a rule that explicitly disables a flag is read as enabling it. The reader must honor the val.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'font-boolean-flag-honors-explicit-false',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A boolean font flag serialized with an explicit-false value (e.g. <b val="0"/>) is read as ' +
    'false, while a bare tag (<b/>) and an explicit-true tag (<b val="1"/>) are read as true — the ' +
    'reader honors the val attribute rather than returning true on tag presence alone.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a bare bold tag reads as true (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {bareTag} = await api.fontExplicitFalseBoldReport();
        assert.strictEqual(bareTag, true, 'a bare <b/> means bold');
      },
    },
    {
      name: 'an explicit-true bold tag reads as true (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {valOne} = await api.fontExplicitFalseBoldReport();
        assert.strictEqual(valOne, true, '<b val="1"/> means bold');
      },
    },
    {
      name: 'an explicit-false bold tag reads as false',
      baseline: 'fail',
      async expect(api, assert) {
        const {valZero} = await api.fontExplicitFalseBoldReport();
        assert.strictEqual(valZero, false, '<b val="0"/> means NOT bold — the val attribute must be honored, not ignored');
      },
    },
    {
      name: 'an explicit-false italic tag reads as not italic',
      baseline: 'fail',
      async expect(api, assert) {
        const {italic} = await api.fontExplicitOffFlagsReport();
        assert.ok(!italic, '<i val="0"/> means NOT italic');
      },
    },
    {
      name: 'an explicit-false strikethrough tag reads as not struck',
      baseline: 'fail',
      async expect(api, assert) {
        const {strike} = await api.fontExplicitOffFlagsReport();
        assert.ok(!strike, '<strike val="0"/> means NOT struck');
      },
    },
    {
      name: 'an underline tag with value "none" reads as not underlined',
      baseline: 'fail',
      async expect(api, assert) {
        const {underline} = await api.fontExplicitOffFlagsReport();
        assert.ok(!underline && underline !== 'none', '<u val="none"/> means NOT underlined — the string "none" is not truthy underline');
      },
    },
  ],
};
