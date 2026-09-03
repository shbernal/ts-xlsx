// Namespace resolution, driven the way a reader drives it: open/close around a real scan.
//
// Every case here is a document the rest of the reader handles perfectly and that a prefix-matching
// path silently loses a feature from. They are not exotic: binding the main namespace to a prefix, or
// the relationships namespace to something other than `r`, is legal and real toolchains do both.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {RELATIONSHIPS_NS} from '../io/opc/namespaces.ts';
import {SPREADSHEETML_NS} from '../io/xlsx/namespaces.ts';
import {NamespaceScope} from './xml-namespaces.ts';
import {type XmlAttributes, xmlEvents} from './xml-scan.ts';

/** Drive a scan through the scope, calling `visit` on every open element. */
function walk(
  xml: string,
  visit: (scope: NamespaceScope, name: string, attrs: XmlAttributes) => void,
) {
  const scope = new NamespaceScope();
  for (const event of xmlEvents(xml)) {
    if (event.kind === 'open') {
      scope.open(event.attrs);
      visit(scope, event.name, event.attrs);
      if (event.selfClosing) scope.close();
    } else if (event.kind === 'close') {
      scope.close();
    }
  }
}

test('an element resolves to its namespace whether it is prefixed or not', () => {
  const seen: (string | undefined)[] = [];
  walk(`<x:worksheet xmlns:x="${SPREADSHEETML_NS}"><x:sheetData/></x:worksheet>`, (scope, name) =>
    seen.push(scope.elementNamespace(name)),
  );
  assert.deepEqual(seen, [SPREADSHEETML_NS, SPREADSHEETML_NS]);

  const unprefixed: (string | undefined)[] = [];
  walk(`<worksheet xmlns="${SPREADSHEETML_NS}"><sheetData/></worksheet>`, (scope, name) =>
    unprefixed.push(scope.elementNamespace(name)),
  );
  assert.deepEqual(unprefixed, [SPREADSHEETML_NS, SPREADSHEETML_NS]);
});

test('an unbound prefix resolves to nothing rather than to the default namespace', () => {
  const seen: (string | undefined)[] = [];
  walk(`<root xmlns="${SPREADSHEETML_NS}"><ext:thing/></root>`, (scope, name) =>
    seen.push(scope.elementNamespace(name)),
  );
  assert.deepEqual(seen, [SPREADSHEETML_NS, undefined]);
});

test('a relationship attribute is found under whatever prefix the file bound', () => {
  for (const prefix of ['r', 'rel', 'q1']) {
    const found: (string | undefined)[] = [];
    walk(
      `<workbook xmlns:${prefix}="${RELATIONSHIPS_NS}"><sheet ${prefix}:id="rId7"/></workbook>`,
      (scope, name, attrs) => {
        if (name === 'sheet') found.push(scope.attr(attrs, RELATIONSHIPS_NS, 'id'));
      },
    );
    assert.deepEqual(found, ['rId7'], `prefix ${prefix}`);
  }
});

test('an unprefixed attribute is in no namespace, so it never matches one', () => {
  // XML says so: a default namespace applies to elements and never to attributes. Every namespaced
  // attribute OOXML writes (`r:id`, `r:embed`, `xml:space`) carries a prefix for exactly that reason.
  const found: (string | undefined)[] = [];
  walk(
    `<workbook xmlns="${RELATIONSHIPS_NS}"><sheet id="rId7"/></workbook>`,
    (scope, name, attrs) => {
      if (name === 'sheet') found.push(scope.attr(attrs, RELATIONSHIPS_NS, 'id'));
    },
  );
  assert.deepEqual(found, [undefined]);
});

test('a declaration binds only within its element, and restores what it shadowed', () => {
  const OTHER = 'urn:other';
  const seen: (string | undefined)[] = [];
  walk(
    `<root xmlns:p="${SPREADSHEETML_NS}">` +
      `<p:outer/>` +
      `<mid xmlns:p="${OTHER}"><p:inner/></mid>` +
      `<p:after/>` +
      `</root>`,
    (scope, name) => {
      if (name.startsWith('p:')) seen.push(scope.elementNamespace(name));
    },
  );
  assert.deepEqual(
    seen,
    [SPREADSHEETML_NS, OTHER, SPREADSHEETML_NS],
    'the inner rebinding applies only inside `mid`',
  );
});

test('a prefix declared on a self-closing element does not leak to its siblings', () => {
  const seen: (string | undefined)[] = [];
  walk(`<root><a xmlns:p="${SPREADSHEETML_NS}"/><p:b/></root>`, (scope, name) => {
    if (name === 'p:b') seen.push(scope.elementNamespace(name));
  });
  assert.deepEqual(seen, [undefined]);
});

test('isElementIn separates a prefixed main namespace from an extension one', () => {
  // The discrimination that `name.includes(':')` was standing in for, and which it got backwards for
  // any file that prefixes the main namespace: every element then has a colon in it.
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const results: {name: string; main: boolean; x14: boolean}[] = [];
  walk(
    `<x:worksheet xmlns:x="${SPREADSHEETML_NS}" xmlns:x14="${X14}">` +
      `<x:dataValidation/><x14:dataValidation/></x:worksheet>`,
    (scope, name) => {
      results.push({
        name,
        main: scope.isElementIn(name, SPREADSHEETML_NS),
        x14: scope.isElementIn(name, X14),
      });
    },
  );
  assert.deepEqual(results, [
    {name: 'x:worksheet', main: true, x14: false},
    {name: 'x:dataValidation', main: true, x14: false},
    {name: 'x14:dataValidation', main: false, x14: true},
  ]);
});
