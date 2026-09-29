# T04b — Filling a new content control with `text.setValue` loses the token's formatting (fixed in 2.15.0)

**Ticket:** T04 · Run-property preservation · **Type:** Bug. Fixed in `@superdoc/sdk` 2.15.0; kept as a regression check.

## Why we do this

Our fill flow wraps a styled placeholder in a content control, then writes the merge value into it with
`contentControls.text.setValue`. The value must inherit the placeholder's formatting, the same way `replace()` does.

## What the script does

Fixture `fixtures/styled-token-body.docx` (synthetic, `build.mjs`). Body paragraph `30000001` reads
`Retention: {{fancy_field}} of the sum.`. The token run carries
`<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:i/><w:color w:val="C00000"/><w:sz w:val="24"/><w:lang w:val="pt-BR"/>`.

Three variants, each on a fresh open, with the token at offsets 11..26:

1. `doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: 'fancy_field', at: <token> })`, then `doc.contentControls.text.setValue({ target, value: 'Fancy Value' })`
2. `doc.create.contentControl(...)` only, with no value written
3. `doc.replace({ target: <token>, text: 'Fancy Value' })` only, as the control

The script compares the `<w:rPr>` of the written text in the exported `word/document.xml` with the input.

## Expected

The value written by `text.setValue` keeps the token's `<w:rPr>`, as `replace()` does.

## Actual

- **`@superdoc/sdk` 2.13.0:** variants 1 and 2 lose the `<w:rPr>`, so the loss happens at `create.contentControl` and not at `setValue`. Variant 3 (`replace()`) keeps it.
- **2.15.0 and 2.16.0-next.17:** all three keep the full `<w:rPr>`, so the script prints `FIXED?`.

Our original observation used a style reference (`<w:rStyle w:val="Strong"/>`) together with direct formatting.
This synthetic fixture has no styles part, so it tests direct formatting only.

See [RESULTS.md](../../RESULTS.md) for every version.

## Done when

It already is, from 2.15.0. We keep the case to catch a regression.

## Our workaround today

On 2.13.0 we capture the run formatting before minting and reapply it after `setValue`.

## Question for SuperDoc

Does `create.contentControl` also preserve `<w:rStyle>` references to character styles from 2.15.0 on, or only direct formatting?
