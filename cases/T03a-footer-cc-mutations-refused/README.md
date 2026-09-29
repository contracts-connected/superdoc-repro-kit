# T03a — Most content-control mutations are refused on a footer control, while `capabilities.resolve` says they are supported

**Ticket:** T03 · Header/footer story support · **Type:** Missing capability + misleading capability report

## Why we do this

Our contracts carry initials, signature dates and merge values in the page footer, as content controls
authored in the template. The server-side fill has to set, clear, relock and finally flatten (unwrap)
those controls exactly as it does in the body. `contentControls.selectByTag` finds a footer control
(its target carries a `headerFooterPart` story), the schema accepts a story target on every mutation,
and from 2.14.0 on `capabilities.resolve` answers `supported` for all of them. Several then throw.

## What the script does

For each of 12 operations, on a fresh open of the fixture:

1. `doc.contentControls.selectByTag({ tag })` and pick the item whose `target.story.storyType` is
   `headerFooterPart`.
2. `doc.capabilities.resolve({ operationId, input })` (when the SDK has it).
3. Call the operation with that target, for example
   `doc.contentControls.clearContent({ target })` or `doc.contentControls.text.setValue({ target, value: 'JD' })`.
4. Save and check the effect in the exported `word/footer1.xml` (content, alias, lock, presence of the control).

| Fixture | Content |
| --- | --- |
| `footer-two-rich-sdts.docx` | one footer paragraph with two inline rich-text controls `seller_initials` and `buyer_initials` (no lock); PAGE field in the sibling paragraph |
| `footer-plain-text-sdt.docx` | a footer with one plain-text control `initials` (for `text.setValue` / `text.clearValue`) |

## Expected

Every mutation behaves on a footer control as it does on a body control; or, where footer support is
not implemented, `capabilities.resolve` reports it as unsupported so a caller can plan without trying.

## Actual

See [RESULTS.md](../../RESULTS.md) for every version and platform.

- **2.13.0:** only `getContent` and `replaceContent` work. The other 10 throw
  `CAPABILITY_UNAVAILABLE Content control mutations outside the document body are not supported by this adapter.`
- **2.15.0 and 2.16.0-next.17:** `patch`, `setLockMode`, `text.setValue` and `unwrap` now work in footers.
  `clearContent`, `appendContent`, `prependContent`, `insertAfter`, `delete` and `text.clearValue` still throw
  the same `CAPABILITY_UNAVAILABLE`, and `capabilities.resolve` answers `supported` for every one of them.

## Done when

The script prints `FIXED?`: every operation takes effect in the exported footer, or `capabilities.resolve`
reports the refused ones as unsupported.

## Our workaround today

Footer merge values are written with `replace()` on the token's text range (see T02b for the guards
that needs); footer controls are never cleared or deleted by our code. Initials and signatures in the
footer stay as text with a visible warning to the user.

## Question for SuperDoc

Is full header/footer parity for content-control mutations planned, and on what timeline? Until then,
can `capabilities.resolve` return an unsupported answer (for example `TARGET_UNSUPPORTED`) for a
header/footer target?
