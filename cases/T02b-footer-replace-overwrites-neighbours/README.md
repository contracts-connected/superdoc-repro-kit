# T02b — `replace()` in a footer overwrites neighbouring text when offsets come from the sourceMap

**Ticket:** T02 · `replace()` target ranges and coordinate spaces · **Type:** Bug (silent data loss)

## Why we do this

We fill merge values (contract numbers, names) inside header and footer stories with `replace()`, because
content-control mutations are not available there (see T03a). We need a token's offsets inside a footer
paragraph, and the Document API gives us three readings of the same paragraph. When a tracked deletion, a
tracked move-from or an `mc:AlternateContent` run precedes the token, the readings disagree. `replace()` accepts
any of them, returns `success:true`, and silently overwrites the wrong characters.

## What the script does

Fixture `fixtures/footer-revisions-and-alternate-content.docx` (synthetic, built by `build.mjs`): one footer with
four paragraphs, each holding one `{{token}}`.

| Paragraph (`w14:paraId`) | What precedes the token |
| --- | --- |
| `1F000001` | nothing special (control) |
| `1F000008` | `<w:del>` containing "deleted " |
| `1F000012` | `<w:moveFrom>` containing "moved " |
| `1F000013` | `<w:r><mc:AlternateContent>` with Choice "X" and Fallback "YYY" |

1. Open the fixture and read the token offsets three ways:
   - `doc.blocks.list({ in: story })` → `textPreview` of the paragraph (same text as `doc.getText({ in: story })`).
   - `doc.projectHtml({ in: story, includeSourceMap: true })` → the token's range mapped back through `sourceMap.entries[].source.segments[].range` (entries report `coordinateSpace: "tracked"`).
   - The part XML: `<w:t>` text of the paragraph, as a caller reading `word/footer1.xml` counts it.
2. For every paragraph × every reading, on a fresh open:
   `doc.replace({ target: { kind: 'selection', start: { kind: 'text', blockId, offset: s, story }, end: { ..., offset: e } }, text: 'VALUE' })`
   with `story = { kind: 'story', storyType: 'headerFooterPart', refId }`.
3. Save, read `word/footer1.xml`, and compare the paragraph with the untouched export where only the token is replaced.

## Expected

One documented coordinate space shared by `getText` / `blocks.list`, the `projectHtml` sourceMap and `replace()`.
A `replace()` whose range does not cover the intended text fails, or can be guarded by an expected-text
precondition, instead of overwriting neighbours.

## Actual

Identical on `@superdoc/sdk` 2.13.0, 2.15.0 and 2.16.0-next.17 (see [RESULTS.md](../../RESULTS.md) for every version):

| Paragraph | `textPreview` offset | sourceMap offset ("tracked") | part XML offset |
| --- | --- | --- | --- |
| `1F000001` control | 1 | 1 | 1 |
| `1F000008` `<w:del>` | 1 | **9** | 1 |
| `1F000012` `<w:moveFrom>` | 1 | **7** | **7** |
| `1F000013` AlternateContent | 2 | 2 | **6** |

`replace()` interprets offsets in the `textPreview` space. Every call below returned `success:true`:

```
1F000008 sourceMap replace  9..23 -> exported "H{{owner_VALUEtail"                         (expected "HVALUEH and a tail")
1F000012 sourceMap replace  7..23 -> exported "moved U{{projVALUE long enough text"        (expected "moved UVALUEU tail long enough text")
1F000013 part XML  replace  6..22 -> exported "VXYYYW{{prVALUEd a long tail of text here"  (expected "VXYYYWVALUEW and a long tail of text here")
```

`getText` / `textPreview` omit the tracked deletion, the tracked move-from and the whole AlternateContent run.
`projectHtml` renders the AlternateContent run as `[unsupported content]`.

## Done when

The script prints `FIXED?`. Either every reading yields the same offsets, or `replace()` rejects a range whose
text is not the text we meant. For example, an `expectedText` precondition that fails instead of writing.

## Our workaround today

Before each `replace()` we check that the paragraph text we planned appears exactly once in the live story text.
After the writes we check that the whole story text equals the old text with only the tokens substituted. If not,
we discard the document. Paragraphs containing `w:moveFrom`, `w:moveTo`, `w:del` or `mc:AlternateContent` are
skipped with a warning, and their tokens stay unfilled.

## Question for SuperDoc

1. Which coordinate space does `replace()` use, and how do we convert a `sourceMap` range (labelled `"tracked"`) into it?
2. Could `replace()` accept an expected-text precondition so a stale or shifted range fails instead of overwriting?
