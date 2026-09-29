# T01a — `create.contentControl` next to a `<w:br/>` in the same run deletes sibling content

**Ticket:** T01 · selection mapping around `<w:br/>` / `<w:tab/>` · **Type:** Bug (silent data loss) on 2.13.0; clean refusal from 2.15.0 on

## Why we do this
Our import converts legacy `{{placeholders}}` in customer templates into inline content controls. Real templates
often put many placeholders in one paragraph, separated only by hard line breaks, frequently inside a single
`<w:r>` (one run with several `<w:t>` nodes and `<w:br/>` between them). Each placeholder becomes its own control.

## What the script does
Two fixtures, one paragraph each, everything in ONE run:

| Fixture | Paragraph | Target |
| --- | --- | --- |
| `one-run-two-tokens-br.docx` | `<w:t>{{f1}}</w:t><w:br/><w:t>{{f2}}</w:t>` (paraId `00000005`) | `{{f1}}`, offsets 0..6 |
| `one-run-prose-br-token.docx` | `<w:t>Line one prose.</w:t><w:br/><w:t>{{t}} more prose.</w:t>` (paraId `0000000A`) | `{{t}}`, offsets 16..21 |

1. Open the fixture with `@superdoc/sdk`.
2. Confirm the target without mutating: `doc.delete({ target, dryRun: true })` resolves exactly the token text.
3. `doc.create.contentControl({ kind: 'inline', controlType: 'text', tag, at: { kind: 'selection', start: { kind: 'text', blockId, offset }, end: { … } } })`
4. Save, read `word/document.xml`, compare the paragraph with the input.

## Expected
The SDT wraps exactly the token; every sibling `<w:t>` and `<w:br/>` survives. If the host cannot do that, the
call fails (throws or returns a failure) and the document is untouched.

## Actual
- **2.13.0:** `success:true` with a node id, and the export has lost content. Shape 1 loses the `<w:br/>` and `{{f2}}`.
  Shape 2 loses `Line one prose.` and the break, and the SDT wraps `{t}} ` (off by one on both edges).
- **2.15.0 and 2.16.0-next.17:** the call throws `CAPABILITY_UNAVAILABLE Content-control selection spans unsupported
  inline structures.` and nothing is lost. The data loss is gone; converting these placeholders is still impossible
  (see T01d for the real template).

See RESULTS.md for every version and platform.

## Done when
The script prints `FIXED?` (no silent loss). It already does on 2.15.0+.

## Our workaround today
We rebuild the paragraph before minting: delete the whole span, reinsert the tokens without breaks, create the
control with nothing after it, reattach the trailing text, restore each break with `insert({ type: 'html', value: '<br>' })`,
then fill with `text.setValue`. It works, but it is slow and fragile.

## Question for SuperDoc
Now that 2.15.0 refuses cleanly: is converting a token that shares a run with a `<w:br/>` (or a `<w:tab/>`) planned?
The refusal leaves real templates unconvertible (T01d).
