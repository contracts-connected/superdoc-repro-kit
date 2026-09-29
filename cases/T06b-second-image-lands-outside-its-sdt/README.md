# T06b — Two initials controls in one footer paragraph: the second image lands outside its control

**Ticket:** T06 — images and offsets inside content controls · **Type:** Bug (silent misplacement)

## Why we do this
A typical work-order footer has both signers' initials in one paragraph: `Subcontractor initials: [seller] <tab> Contractor initials: [buyer]`. The seller usually signs first.

SuperDoc suggested this recipe for placing signer initials in a footer. We run it verbatim, except `size.unit`, which the SDK rejects (see [T06c](../T06c-images-list-ignores-in-and-size-unit/README.md)):

1. `contentControls.selectByTag({ tag })`, keeping the item whose target story is `headerFooterPart`.
2. `contentControls.replaceContent({ target, content: ' ', format: 'text' })`.
3. Re-resolve the control with `selectByTag` again.
4. `create.image({ src, alt, size: { width: 72, height: 28 }, in: story, at: { kind: 'inParagraph', target: { kind: 'block', nodeType: 'paragraph', nodeId: selectionTarget.end.blockId, story }, offset: selectionTarget.end.offset } })`.

## What the script does
The fixture has one footer paragraph with two inline rich-text SDTs, `seller_initials` and `buyer_initials`, both `sdtLocked` as in our templates. The `PAGE` field sits in the sibling paragraph, so it cannot interfere (see [T06a](../T06a-create-image-fails-beside-page-field/README.md)).

The script runs the recipe on a fresh open, first seller then buyer, and again buyer then seller. It walks the exported paragraph and prints where each `<w:drawing>` sits, counting an existing image as 0 and as 1 character.

## Expected
In both orders, each `create.image` returns `success:true` and each drawing is inside its own tagged `<w:sdtContent>`.

## Actual
**Seller, then buyer:** both calls return `success:true`, but the buyer drawing is exported just before the buyer SDT, outside any control.

- `selectByTag` reports the buyer control ending at offset 39. Offset 39 is inside the control only if the seller image counts as zero characters.
- `create.image` counts that image as one character, so 39 lands one position earlier, outside the control.

**Buyer, then seller:** both drawings are inside their own controls.

See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
The script prints `FIXED?`, meaning both drawings are inside their controls in both orders.

## Our workaround today
None that is safe. Signing order is decided by the parties, not by us.

## Question for SuperDoc
Which offsets are authoritative when a paragraph already holds an inline image? Could `create.image` accept the content control itself as its target, for example `at: { kind: 'contentControl', target, position: 'end' }`, so callers never compute offsets?
