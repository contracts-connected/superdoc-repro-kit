# T08a — `contentControls.date.setValue` stores the date but writes no visible text

**Ticket:** T08 · Content-control values and identity · **Type:** Bug

## Why we do this
Signature ceremonies stamp dates such as `[buyer_signature_date]` into the contract. We convert the token into a date content control and set its value. The signer-facing document must show the date. Today it keeps showing the raw token.

## What the script does
1. Builds a paragraph `Signed on: {{sig_date}}` with the token at offsets 11..23.
2. Creates an inline `date` control over the token, calls `contentControls.date.setValue({ target, value: '2026-09-21' })`, and saves.
3. Control case: an inline `text` control over the same token, then `contentControls.text.setValue` with the same string.
4. Reads `<w:sdtContent>` and `<w:date w:fullDate>` from the exported `word/document.xml`.

```bash
node cases/T08a-date-setvalue-writes-no-visible-text/repro.mjs 2.15.0
```

## Expected
After `date.setValue`, the control's content shows the date, the way `text.setValue` replaces a text control's content.

## Actual
`date.setValue` returns success and writes `w:fullDate="2026-09-21"`, but `<w:sdtContent>` still contains `{{sig_date}}`. The text control shows `2026-09-21`. See [RESULTS.md](../../RESULTS.md).

## Done when
The script prints `FIXED?`: the exported date control shows the date in its content.

## Our workaround today
We cannot write the visible text ourselves. After `date.setValue` the control's selection is zero-width, so formatting and text cannot be applied to it. We abandon the conversion, and the raw token stays in the document with a warning.

## Question for SuperDoc
Is the empty body intended? If so, what is the supported way to set a date control's display text, including a display format such as `w:dateFormat`?
