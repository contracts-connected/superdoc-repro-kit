# T04a — A new content control drops the wrapped run's formatting (fixed in 2.15.0)

**Ticket:** T04 · Run-property preservation · **Type:** Bug. Fixed in `@superdoc/sdk` 2.15.0; kept as a regression check.

## Why we do this

We convert legacy placeholders such as `[buyer_initials]` into inline rich-text content controls. The placeholder
text is styled by the template author, for example Arial, bold, dark blue, 10pt. The control must keep that
styling, or the signed contract changes appearance.

## What the script does

Fixture `fixtures/styled-tokens-footer-and-body.docx` (synthetic, `build.mjs`). Every run carries
`<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="0B5394"/><w:sz w:val="20"/>`.

| Shape | Where | Paragraph |
| --- | --- | --- |
| token shares a run with its label: `Initials: [buyer_initials]` | footer | `50000001` |
| token in its own run: `Seller: ` + `[seller_initials]` | footer | `50000002` |
| same-run token: `Body initials: [body_initials]` | body (comparison) | `10000001` |

For each shape, on a fresh open:

```js
doc.create.contentControl({ kind: 'inline', controlType: 'richText', tag, alias: tag,
  at: { kind: 'selection', start: { kind: 'text', blockId, offset: s, story }, end: { kind: 'text', blockId, offset: e, story } } })
```

The script then compares the `<w:rPr>` of the run inside the exported `<w:sdtContent>` with the token run's `<w:rPr>` in the input.

## Expected

Every mint succeeds, the SDT wraps exactly the token, and the run inside it keeps the original `<w:rPr>`.

## Actual

- **`@superdoc/sdk` 2.13.0:** all three shapes, footer and body, are created with `success:true`, but the run inside the SDT has **no `<w:rPr>`** (`<w:r><w:t>[buyer_initials]</w:t></w:r>`).
- **2.15.0 and 2.16.0-next.17:** all three keep the full `<w:rPr>`, so the script prints `FIXED?`.

See [RESULTS.md](../../RESULTS.md) for every version and platform.

## Done when

It already is, from 2.15.0. We keep the case to catch a regression.

## Our workaround today

On 2.13.0 we read the run formatting before minting and reapply it with `format.*` afterwards. We will remove this
when we move to 2.15.0.

## Question for SuperDoc

None, beyond confirming that the 2.15.0 behaviour is the intended contract.
