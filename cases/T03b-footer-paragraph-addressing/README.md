# T03b — `find()` returns nothing in header/footer stories, and `blocks.list` does not enumerate table-cell paragraphs

**Ticket:** T03 · Header/footer story support · **Type:** Bug / contract question

## Why we do this

Our purchase-order footers are tables. Merge tokens (`{{scope_contract_number}}`, `[seller_initials]`) live in
table-cell paragraphs, sometimes inside a locked block SDT. To write a value we need a deterministic address for
the paragraph that holds the token. Today we use `blockId` = the paragraph's `w14:paraId`, plus offsets. We want
to know whether that is a supported contract, and why `find()` does not see footer text.

## What the script does

Two inputs:

- `fixtures/footer-plain-and-table.docx` (synthetic, `build.mjs`). The first-page footer is a plain paragraph `50000001` holding `Initials: [buyer_initials]`. The default footer is a 2×2 table whose cell paragraphs are `61000001`…`61000004`, with cell 1 holding `No: {{contract_number}}`.
- `fixtures/sanitized/po-locked-footer.docx`, a real purchase order, sanitized. It has two footers built as tables, one inside a `contentLocked` block SDT.

For each footer part (`story = { kind: 'story', storyType: 'headerFooterPart', refId }`):

1. `doc.getText({ in: story })`, which does contain the token.
2. `doc.query.match({ in: story, select: { type: 'text', pattern: token, mode: 'contains' } })`
3. `doc.find({ in: story, select: { type: 'text', pattern: token, mode: 'contains' } })`
4. `doc.query.match({ in: { kind: 'story', storyType: 'headerFooterSlot', section: { kind: 'section', sectionId }, headerFooterKind: 'footer', variant }, select: … })`
5. `doc.blocks.list({ in: story })`
6. `doc.replace({ target: <selection in the paragraph whose blockId = w14:paraId>, text: 'VALUE' })`, then check the export.

As a control, `doc.find({ in: { kind: 'story', storyType: 'body' }, select: … })` on the body returns the match (`total: 1`).

## Expected

Text that `getText()` returns for a footer story is found by `find()` scoped to that story, as it is in the body.
`blocks.list` gives an address for every paragraph, including table cells. The paragraph address used for writes is documented.

## Actual

Identical on `@superdoc/sdk` 2.13.0, 2.15.0 and 2.16.0-next.17 (see [RESULTS.md](../../RESULTS.md)):

```
getText({ in: headerFooterPart })                   contains token: true
query.match({ in: headerFooterPart, … })            -> 1          works
find({ in: headerFooterPart, … })                   -> 0          (body control: 1)
query.match({ in: headerFooterSlot section-0/… })   -> 1          works
blocks.list({ in: headerFooterPart })               -> "table tbl:61000001, paragraph 61000009"   cell paragraphs not listed
replace() with blockId = w14:paraId 61000001        -> success:true, exported "No: VALUE"   works
```

The sanitized purchase order gives the same results: its 13 table-cell paragraphs per footer are not listed.

In earlier runs `query.match` in a header/footer story sometimes returned 0. This script finds it working on
every version above; run-to-run variation of `query.match` is covered by its own case.

## Done when

The script prints `FIXED?`. `find()` returns footer matches, and `blocks.list` enumerates table-cell paragraphs,
or its docs say how to reach them, for example a `tables.getCells` equivalent for header/footer stories.

## Our workaround today

We read the footer part XML ourselves, compute offsets from `<w:t>` text, and address the paragraph by its
`w14:paraId`. This works, but it depends on an identity we have not seen documented.

## Question for SuperDoc

1. Is `blockId === w14:paraId` for header/footer paragraphs, including table-cell paragraphs, a supported, stable contract?
2. Is `find()` meant to search header/footer stories?
3. What is the supported way to enumerate paragraphs inside a table in a header/footer story?
