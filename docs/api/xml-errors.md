# Xml Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `XmlParseError`

<sub>class</sub>

Thrown when XML text cannot be scanned into parse events — an unterminated tag, comment, CDATA
section, markup declaration or processing instruction.

This is a *typed* failure rather than the native `SyntaxError` it reads like, because the text
being parsed is almost always a part lifted out of an untrusted `.xlsx`. A caller wrapping
`readXlsx` needs to tell "the file I was handed is corrupt" from "something in my own code threw a
`SyntaxError`", and a native error gives them no way to. The message still names the construct that
did not terminate; only the type changed.

```ts
class XmlParseError extends XlsxError {
  override readonly name = 'XmlParseError';
  override readonly code = 'malformed-input';
}
```
