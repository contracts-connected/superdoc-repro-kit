# T01b — `create.contentControl` next to a `<w:tab/>` in the same run deletes the tab and the next token

**Ticket:** T01 · selection mapping around `<w:br/>` / `<w:tab/>` · **Type:** Bug (silent data loss) on 2.13.0; clean refusal from 2.15.0 on

## Why we do this
Same import as T01a. Templates also separate placeholders with tabs (labels and values aligned on a tab stop).
Here the dangerous case is the placeholder *before* the tab: it looks safe, but minting it destroys what follows.

## What the script does
| Fixture | Paragraph (ONE run) | Target |
| --- | --- | --- |
| `one-run-two-tokens-tab.docx` | `<w:t>{{f1}}</w:t><w:tab/><w:t>{{f2}}</w:t>` (paraId `00000006`) | `{{f1}}` (before the tab), offsets 0..6 |
| `one-run-prose-tab-token.docx` | `<w:t>Line one prose.</w:t><w:tab/><w:t>{{t}} more prose.</w:t>` (paraId `0000000D`) | `{{t}}` (after the tab), offsets 16..21 |

1. Open, confirm the range with `doc.delete({ target, dryRun: true })`.
2. `doc.create.contentControl({ kind: 'inline', controlType: 'text', tag, at: <selection> })`
3. Save and compare the exported paragraph with the input.

## Expected
The SDT wraps exactly the token; the `<w:tab/>` and every sibling text node survive, or the call fails and nothing changes.

## Actual
- **2.13.0:** `success:true`. Token before the tab: the tab and `{{f2}}` are gone. Token after the tab: the prose before it
  and the tab are gone, and the SDT wraps `{t}} `.
- **2.15.0 and 2.16.0-next.17:** `CAPABILITY_UNAVAILABLE Content-control selection spans unsupported inline structures.`,
  nothing lost.

See RESULTS.md for every version and platform.

## Done when
The script prints `FIXED?` (it does on 2.15.0+).

## Our workaround today
We cannot rebuild around a tab: no Document API primitive re-inserts a literal `<w:tab/>` (see T02c). Any group of
placeholders involving a tab is left unconverted, with a warning shown to the user.

## Question for SuperDoc
Same as T01a: is converting a token that shares a run with a tab planned?
