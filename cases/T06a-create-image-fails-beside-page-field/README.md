# T06a — `create.image` in a footer paragraph that contains a PAGE field fails, while `dryRun` says it succeeds

**Ticket:** T06 — images and offsets inside content controls · **Type:** Bug

## Why we do this
Signers' initials go into the page footer as an image inside an inline rich-text content control. Real footers usually put a page number field in the same paragraph as the initials, for example `Initials: [ ]  Page 3`.

SuperDoc suggested this recipe for placing signer initials in a footer. We run it verbatim, except `size.unit`, which the SDK rejects (see [T06c](../T06c-images-list-ignores-in-and-size-unit/README.md)):

1. `contentControls.selectByTag({ tag })`, keeping the item whose target story is `headerFooterPart`.
2. `contentControls.replaceContent({ target, content: ' ', format: 'text' })`.
3. Re-resolve the control with `selectByTag` again.
4. `create.image({ src, alt, size: { width: 72, height: 28 }, in: story, at: { kind: 'inParagraph', target: { kind: 'block', nodeType: 'paragraph', nodeId: selectionTarget.end.blockId, story }, offset: selectionTarget.end.offset } })`.

## What the script does
It runs the recipe on two footers that differ only in where the `PAGE` complex field (`w:fldChar`) sits:

- **Same paragraph:** `Initials: ` + SDT `buyer_initials` + tab + `Page {PAGE}`.
- **Sibling paragraph (control):** the SDT alone, with `Page {PAGE}` in the next paragraph.

Right before the real `create.image`, it asks `capabilities.resolve` and `create.image({ ...input, dryRun: true })` about the very same input. It then checks the exported `word/footer1.xml` for a `<w:drawing>` inside the SDT.

## Expected
The image lands inside the SDT in both layouts. If the field paragraph is unsupported, `dryRun` and `capabilities.resolve` report that before the real call.

## Actual
- **Same paragraph:** `capabilities.resolve` returns `supported` and `dryRun` returns `success:true`, but the real call throws `TARGET_NOT_FOUND image-paragraph-not-found`. No drawing is exported.
- **Sibling paragraph:** `success:true`, and the drawing is inside the SDT.

We also saw the same result in real Chromium on `superdoc@2.18.0-next.8` and `next.9`. See [RESULTS.md](../../RESULTS.md) for every SDK version.

## Done when
The script prints `FIXED?`. That means the image lands inside the SDT when the paragraph also holds a `PAGE` field.

## Our workaround today
We use the recipe only when no field shares the paragraph. Real templates with that layout keep their initials as text.

## Question for SuperDoc
Is a fix planned for paragraphs that contain complex fields? Until then, can `dryRun` and `capabilities.resolve` report the failure?
