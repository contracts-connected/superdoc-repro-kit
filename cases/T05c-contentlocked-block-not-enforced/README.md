# T05c — `contentLocked` on a block content control is not enforced on its contents

**Ticket:** T05 — protection, permission ranges and locks · **Type:** Contract question

## Why we do this
Google Docs exports wrap some headers and footers in a block SDT tagged `goog_rdk_0` with `<w:lock w:val="contentLocked"/>`. Our templates arrive in that shape. We fill merge values and mint signer fields inside such a wrapper today. That works only because the lock is not enforced. We need to know whether we may rely on this.

## What the script does
The fixture's first-page footer (`titlePg`) is a block SDT `goog_rdk_0` with `contentLocked`. It holds two paragraphs:

- `Seller initials: [seller_initials]`, where the token is raw text;
- `Buyer initials: ` followed by an inline rich-text SDT `buyer_initials`.

The script does two things inside the locked block:

- **(a)** `create.contentControl` over the raw `[seller_initials]` token;
- **(b)** `contentControls.replaceContent` on the inline `buyer_initials` SDT.

It then checks both effects in the exported `word/footer1.xml`.

## Expected
Either edits inside a `contentLocked` block are refused with a lock error, or the documentation states that the Document API does not enforce `contentLocked` on contents.

## Actual
`selectByTag('goog_rdk_0')` reports `lockMode: contentLocked`. Both edits return `success:true` and both are present in the exported XML. The wrapper keeps its `contentLocked` lock.

By contrast, a locked inline SDT does refuse `replace()` with `text-range-in-sdt`. See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
The contract is written down. If you decide to enforce the lock, this script will print `FIXED?`, and we will then need a documented way to fill such templates.

## Our workaround today
None. We depend on today's behaviour.

## Question for SuperDoc
Is `contentLocked` on a block SDT meant to be enforced against Document API writes inside it? If yes, what is the supported way to fill a template that arrives with this wrapper?
