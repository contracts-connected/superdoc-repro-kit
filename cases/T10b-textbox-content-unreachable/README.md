# T10b — Text inside a text box: how to read and write it

**Ticket:** T10 · Read paths · **Type:** Contract question (the write path works on current versions)

## Why we do this
Contract templates from our previous version place merge fields such as `{{project_name}}` inside text boxes (`wps:txbx`). We fill them server-side. On earlier SDK versions nothing could read or write text-box content through the Document API, so we rewrite the part XML directly.

## What the script does
1. Builds a body paragraph, a paragraph holding an inline drawing whose text box contains "Textbox content {{project_name}}", and another body paragraph.
2. Reads it with `getText`, `blocks.list`, `query.match` with `in: body`, and `query.match` without `in`.
3. Writes with `replace()` using the target that `query.match` returned, and with `replace()` addressing the text-box paragraph's `w14:paraId` directly.
4. Checks the exported text-box content.

```bash
node cases/T10b-textbox-content-unreachable/repro.mjs 2.15.0
```

## Expected
Text-box content can be found and written through the Document API.

## Actual
On the versions we ran, `query.match` without `in` returns the text with a story of `{ storyType: 'textbox', textboxId: 'tb0' }`. `replace()` on that target writes into the text box, so the script prints `FIXED?`.

`getText({})` and `blocks.list({})` do not include text-box content. `query.match` with `in: body` returns 0. Addressing the text-box paragraph by its `w14:paraId` throws `TARGET_NOT_FOUND`. See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
It already passes for the write path. We would like the discovery side confirmed (see the questions).

## Our workaround today
Raw part-XML rewrite of `wps:txbx` content after save. We plan to move to the `textbox` story once the questions below are answered.

## Question for SuperDoc
1. Is `{ kind: 'story', storyType: 'textbox', textboxId }` a supported, documented story? From which versions?
2. How do we enumerate a document's text boxes, and their ids, without a text search? `blocks.list` and `getText` skip them.
3. Are `textboxId` values stable across open, save and reopen?
