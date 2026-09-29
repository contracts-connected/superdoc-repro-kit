# T08b — A paragraph's id changes on a block content-control mint and again on `clearContent`

**Ticket:** T08 · Content-control values and identity · **Type:** Bug

## Why we do this
A signature field is a block rich-text content control wrapping one paragraph. To sign, we clear the control and insert the signature image into its inner paragraph with `create.image` `inParagraph`. That needs the inner paragraph's id, and the id changes twice without any receipt reporting the new one.

## What the script does
1. `blocks.list`: the paragraph "buyer signature placeholder" is `00000002`, its `w14:paraId`.
2. `create.contentControl({ kind: 'block', controlType: 'richText', … })` over that paragraph, then `blocks.list` again.
3. `contentControls.clearContent` on the new control, then `blocks.list` again.
4. Prints what `capabilities.get` and `capabilities.resolve` say about `create.image`. Then calls `create.image` `inParagraph` twice: once with the original id and once with the current id.
5. Saves and checks the exported inner paragraph id and whether the drawing sits inside the SDT.

```bash
node cases/T08b-paragraph-id-changes-after-block-mint-and-clear/repro.mjs 2.15.0
```

## Expected
The paragraph keeps its id through the mint and `clearContent`. Failing that, the `create.contentControl` and `clearContent` receipts return the new inner paragraph id.

## Actual
The id goes `00000002` → `00000004` after the mint and → `00000005` after `clearContent`. Neither receipt contains the new id. `create.image` with the original id throws `TARGET_NOT_FOUND`; with the re-listed id it succeeds. See [RESULTS.md](../../RESULTS.md).

On 2.15.0, `capabilities.get` and `capabilities.resolve` now report `create.image` as available, which matches what happens. Earlier versions reported it as unavailable in this situation.

## Done when
The script prints `FIXED?`: the id is stable, or a receipt carries the new id.

## Our workaround today
After every mint or clear, we re-list blocks and pick the paragraph by ordinal. That works as long as nothing else changes the document between the two calls.

## Question for SuperDoc
Is paragraph identity expected to survive block wrapping and `clearContent`? If not, can the receipt include the inner paragraph's new `nodeId`? Or can `create.image` target the content control directly, as in "the control's content"?
