# T01c — Mint an inline content control over a footer token that follows a `<w:tab/>`

**Ticket:** T01 · selection mapping around `<w:br/>` / `<w:tab/>` · **Type:** Missing capability on 2.13.0; fixed in 2.15.0

## Why we do this
Our contract footers read `Initials:` · tab · `[buyer_initials]`. To let a signer place initials there (your recipe:
`selectByTag` → `replaceContent(' ')` → `create.image`), the raw token first has to become an inline rich-text content
control. On 2.13.0 that mint is refused, so a legacy footer placeholder can never become a native field and our
customers see raw brackets in signed documents.

## What the script does
| Fixture | Footer paragraph |
| --- | --- |
| `footer-token-after-tab.docx` | `Initials:` · `<w:tab/>` · `[buyer_initials]` (three runs, rPr Arial / bold / 0B5394 / sz 20) |
| `footer-two-labels-tab-stop.docx` | paragraph tab stop; `Subcontractor Initials: [seller_initials]` · spaces · `<w:tab/>` · `Contractor Initials: [buyer_initials]` · ` tail` |
| `work-order-footer-initials.docx` (sanitized real template, `fixtures/sanitized/`) | our real work-order footer, same layout as the second fixture |

1. Find the footer part: `doc.headerFooters.parts.list({})`, story `{ kind: 'story', storyType: 'headerFooterPart', refId }`.
2. Offsets from the host: the line of `doc.getText({ in: story })` that contains the token.
3. `doc.capabilities.resolve({ operationId: 'create.contentControl', input })` (when the SDK has it), then
   `doc.create.contentControl({ kind: 'inline', controlType: 'richText', tag: 'buyer_initials', alias: 'buyer_initials', at: { kind: 'selection', start: { kind: 'text', blockId: <w14:paraId>, offset, story }, end: { … } } })`
4. Repeat on a fresh open with offsets counted from the part XML (`w:t` text + each run-level tab as one character), for comparison.
5. Save and read the SDT in `word/footer1.xml`: what it wraps, and the rPr inside it.

## Expected
The mint succeeds, and the new SDT wraps exactly `[buyer_initials]`.

## Actual
- **2.13.0:** `CAPABILITY_UNAVAILABLE Content-control selection spans unsupported inline structures.` on both synthetic fixtures.
- **2.15.0 and 2.16.0-next.17:** the mint succeeds, wraps exactly the token, and keeps the run's rPr. Host offsets and
  XML offsets agree. The one-character skew we saw earlier on 2.14.0 (the SDT wrapping `buyer_initials] `) does not
  appear on these versions.
- If the whole package is rejected (`INVALID_INPUT Cannot read content-control identities from malformed package XML.`),
  the script prints a NOTE and does not count it here; that is T07a.

See RESULTS.md for every version and platform.

## Done when
The script prints `FIXED?`. It does on 2.15.0+, so this is a confirmation, not a request.

## Our workaround today
On 2.13.0 we leave the footer token as raw text and show the user a warning.

## Question for SuperDoc
Can you confirm this fix is intended and permanent in the 2.15 line? We plan to move production to SDK 2.15.0 for it.
