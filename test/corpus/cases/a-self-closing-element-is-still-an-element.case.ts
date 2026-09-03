// Cluster: security
//
// Real-world scenario: `<definedName name="Empty"/>` is legal markup. A self-closing element fires no
// close event, and a reader whose text capture commits on the close therefore never commits it: the
// defined name was dropped from the model entirely, and the latch it left set went on absorbing
// whatever character data came next until some unrelated open happened to reset it.
//
// That last part is the reason this is a security cluster rather than a fidelity one. `TextCapture`
// was written to make the hazard structural and says so: what kept the open-coded versions from
// corrupting anything "was the order the next open happened to reset things in, which is an accident
// rather than a property anyone chose, on a path that reads untrusted input". Three readers still
// hand-rolled it.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-self-closing-element-is-still-an-element',
  provenance: {source: 'text-capture-audit'},
  cluster: 'security',
  description:
    'A self-closing element that a reader captures text from is read as the empty element it is, ' +
    'rather than dropped while leaving the capture latched on the text that follows it.',

  behavior: [
    {
      name: 'a self-closing <definedName/> is read, and does not swallow the next name',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.selfClosingDefinedNameReport().names, [
          {name: 'Empty', refersTo: ''},
          {name: 'Later', refersTo: 'S!$A$1'},
        ]);
      },
    },
  ],
} satisfies Case;
