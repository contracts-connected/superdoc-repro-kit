# T06c — The SDK rejects `size.unit`; `images.list({ in })` ignores `in`

**Ticket:** T06 — images and offsets inside content controls · **Type:** Bug (two small contract issues)

## Why we do this
- **`size.unit`:** SuperDoc's initials recipe passes `size: { width, height, unit: 'px' }`. The browser package accepts it, but our server-side host does not.
- **`images.list`:** we use `images.list({ in: story })` to check whether a footer already holds a signer's image before placing another.

## What the script does
- **(a)** Runs the recipe on a footer rich-text SDT with `size: { width: 72, height: 28, unit: 'px' }`.
- **(b)** Inserts one image in body paragraph `10000001`, and none in the footer. It then calls `images.list({})` and `images.list({ in: <footer headerFooterPart story> })`, and checks the exported package to confirm the drawing is only in `word/document.xml`.

## Expected
- **(a)** `size.unit` is accepted, as in the browser package, or the difference is documented.
- **(b)** `images.list({ in: footer })` returns 0 items when the only image is in the body.

## Actual
- **(a)** `VALIDATION_ERROR create image:size.unit is not allowed by schema.` Yet `images.list` itself returns `size: { width, height, unit: 'px' }`.
- **(b)** `images.list({ in: footer })` returns total 1, and the item is the body image `img:1:word_document.xml:rId1`.

See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
The script prints `FIXED?`, meaning `size.unit` is accepted and the footer-scoped list is empty.

## Our workaround today
We omit `size.unit` on the SDK. We filter `images.list` results by the part name embedded in `sdImageId`, which is not a documented contract.

## Question for SuperDoc
Which is the contract for `size.unit` on `create.image`? Should `images.list` honour `in`?
