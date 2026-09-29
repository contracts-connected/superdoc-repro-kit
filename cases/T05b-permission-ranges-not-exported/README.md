# T05b — Permission ranges and the editing restriction after export and reopen

**Ticket:** T05 — protection, permission ranges and locks · **Type:** Bug (regression check)

## Why we do this
A signing round is saved and reopened by the other party. The restriction and each signer's range must survive that round trip, or the second signer opens an unprotected document.

## What the script does
1. Mints a block rich-text signature control.
2. Turns on `protection.setEditingRestriction({ mode: 'readOnly' })`.
3. Creates one permission range with principal `everyone` over the control's text.
4. Saves the document.
5. Searches every XML part of the exported package for `w:permStart`, `w:permEnd` and `w:documentProtection`.
6. Reopens the exported file and calls `protection.get()` and `permissionRanges.list()`.

## Expected
- The range is written as `w:permStart` and `w:permEnd`.
- The restriction is written as `w:documentProtection`.
- After reopening, `protection.get()` reports `readOnly` and `permissionRanges.list()` returns the range.

## Actual
We first saw this on the browser package, `superdoc@2.16.0`. There the range was lost on export: `permissionRanges.list()` returned total 0 after reopening, and the exported document had no `w:permStart`.

On the SDK the round trip works on the versions in [RESULTS.md](../../RESULTS.md) that print `FIXED?`. We keep the case as a regression check. Please confirm the browser package matches.

## Done when
The script prints `FIXED?` on every SDK version, and the same holds for the browser package.

## Our workaround today
None needed on SDK versions where this passes. On the browser host, our application re-grants the ranges after opening a saved round.

## Question for SuperDoc
Is the browser package's export expected to write `w:permStart` and `w:permEnd` the same way the SDK does, and from which `superdoc` version?
