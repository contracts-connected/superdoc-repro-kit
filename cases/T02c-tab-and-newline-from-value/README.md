# T02c — A line break inside a merge value cannot be written through `replace()` or `insert({ type: 'text' })`

**Ticket:** T02 · `replace()` target ranges and coordinate spaces · **Type:** Missing capability / contract question

## Why we do this

Merge values are plain strings from our database. Addresses are multi-line (`"12 Main St\nSpringfield"`), and some
fields are "Label<TAB>Value" pairs. We write them with `replace()` over the placeholder's range. We need a
supported way to turn `\n` into `<w:br/>`, and to know that `\t` becomes `<w:tab/>`.

## What the script does

Fixture `fixtures/body-one-paragraph.docx` (synthetic, `build.mjs`): paragraph `20000001` reads `Address: X.`;
the `X` (offsets 9..10) is the value slot. Each call runs on a fresh open, then the script inspects the exported
`word/document.xml`:

1. `doc.replace({ target: <X>, text: 'A\tB' })`
2. `doc.replace({ target: <X>, text: 'A\nB' })`
3. `doc.insert({ target: <collapsed selection after X>, type: 'text', value: 'A\tB' })`
4. `doc.insert({ target: <after X>, type: 'text', value: 'A\nB' })`
5. `doc.insert({ target: <after X>, type: 'html', value: 'A&#9;B' })`
6. `doc.insert({ target: <after X>, type: 'html', value: 'A<span style="white-space:pre">\t</span>B' })`
7. `doc.insert({ target: <after X>, type: 'html', value: 'A<br>B' })`, the line-break path that works

## Expected

A `\t` in a text value becomes `<w:tab/>` and a `\n` becomes `<w:br/>` through both `replace({ text })` and
`insert({ type: 'text' })`. Failing that, the docs name the supported way to write a line break from a value.

## Actual

Identical on `@superdoc/sdk` 2.13.0, 2.15.0 and 2.16.0-next.17 (see [RESULTS.md](../../RESULTS.md)):

```
replace({ text: "A\tB" })                   success:true   exported "Address: A<TAB>B."   real <w:tab/>
replace({ text: "A\nB" })                   COMMAND_FAILED text-payload-unsupported-control-char
insert({ type: "text", value: "A\tB" })     success:true   exported "Address: XA<TAB>B."  real <w:tab/>
insert({ type: "text", value: "A\nB" })     COMMAND_FAILED text-payload-unsupported-control-char
insert({ type: "html", value: "A&#9;B" })   success:true   exported "Address: XA B."      tab silently became a space
insert({ type: "html", value: "A<br>B" })   success:true   <w:br/> written
```

Tabs work. Line breaks are refused with a clear error through both text paths. The only working route is
`insert({ type: 'html' })`, which cannot replace a range. A tab passed as HTML (`&#9;` or inside
`white-space:pre`) silently becomes a space, with `success:true`.

## Done when

The script prints `FIXED?`, meaning a `\n` in a text value becomes `<w:br/>` through `replace()` and `insert({ type: 'text' })`.

## Our workaround today

When we fill header/footer values, a value that contains a line break is skipped with a warning and stays unfilled.
When our template import has to put a line break back, it uses `insert({ type: 'html', value: '<br>' })` at a
collapsed position. Leading and trailing spaces of an HTML payload are collapsed, so we guard them with a
one-character anchor that we delete afterwards.

## Question for SuperDoc

1. Is there a supported way to write a line break from a value string into a range, for example a `replace()` option or a structured `content` payload?
2. Is the tab-to-space conversion in `insert({ type: 'html' })` intended?
