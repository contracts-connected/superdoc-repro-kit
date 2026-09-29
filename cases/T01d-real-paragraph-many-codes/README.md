# T01d — A real paragraph with 110 placeholders separated by `<w:br/>` cannot be converted

**Ticket:** T01 · selection mapping around `<w:br/>` / `<w:tab/>` · **Type:** Bug (silent data loss) on 2.13.0; Missing capability on 2.15.0+

## Why we do this
This is the paragraph that broke our first live import. It comes from a real field-catalog template: 110 `{{placeholders}}`
in ONE run, separated by `<w:br/>` (usually two breaks between tokens) with a leading break. Our import must turn every
placeholder into an inline content control.

## What the script does
Fixture: `fixtures/sanitized/paragraph-116-codes.docx` (paragraph `00000002`: 110 tokens, 219 `<w:br/>`). Three passes,
each on a fresh open, with offsets counted from the paragraph XML (`w:t` text, each `<w:br/>` = one character):
1. **first-token** — `create.contentControl` on `{{buying_company_name}}` only (range confirmed with `doc.delete({ target, dryRun: true })`).
2. **last-token** — `create.contentControl` on the last token only.
3. **import-all** — every token, last to first, `doc.create.contentControl({ kind: 'inline', controlType: 'text', tag, at })`.

After each pass the script saves and counts, in the exported paragraph, the tokens and breaks that survive, the SDTs,
and any SDT that wraps something other than a whole token.

## Expected
All 110 placeholders become content controls, each wrapping exactly its token, with every token and every `<w:br/>` preserved.

## Actual
- **2.13.0:** first-token and last-token refuse, but **import-all creates 1 control and the export keeps 1 of 110 tokens
  and 0 of 219 breaks, with `success:true`.** This is the silent loss we hit in production (109 placeholders gone).
- **2.15.0 and 2.16.0-next.17:** every call throws `CAPABILITY_UNAVAILABLE Content-control selection spans unsupported
  inline structures.`. Nothing is lost, and nothing can be converted (0 of 110).

See RESULTS.md for every version and platform.

## Done when
import-all creates 110 controls, each wrapping exactly its token, with 110/110 tokens and 219/219 breaks in the export.

## Our workaround today
We rebuild the paragraph ourselves: delete each cluster, reinsert the tokens without breaks, mint, restore the breaks,
fill by node id (details in T01a). It converts all 110 without loss, but it takes about 40 seconds in the browser and
depends on several other host behaviours staying the same.

## Question for SuperDoc
Is minting a control over a token that shares a run with `<w:br/>` planned? This is the most common shape in our
customers' legacy templates.
