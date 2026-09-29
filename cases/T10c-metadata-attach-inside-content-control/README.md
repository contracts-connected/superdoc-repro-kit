# T10c — `metadata.attach` refuses a range inside a content control

**Ticket:** T10 · Read paths · **Type:** Missing capability

## Why we do this
Signature, initials and date fields are content controls. We need to attach our own identity to each field (field id, signer role, signing round) so it survives export and reopen. Anchored metadata looks like the right tool, and the natural anchor is the field's own content.

## What the script does
1. Control case: `metadata.attach` on plain body text (paragraph `00000001`, offsets 0..5).
2. `contentControls.selectByTag({ tag: 'buyer_initials' })`, then `metadata.attach` on that control's `selectionTarget`.
3. Prints `metadata.list` for the namespace.

```bash
node cases/T10c-metadata-attach-inside-content-control/repro.mjs 2.15.0
```

## Expected
`metadata.attach` accepts a range inside a content control, as it does on plain text.

## Actual
Plain text succeeds. Inside the control it throws `COMMAND_FAILED metadata.attach target overlaps unsupported content (text-range-in-sdt).` See [RESULTS.md](../../RESULTS.md).

## Done when
The script prints `FIXED?`: the attach inside the control succeeds and is listed.

## Our workaround today
Field identity lives only in the control's `w:tag` and `w:alias`, which limits what we can store and forces us to encode it in strings.

## Question for SuperDoc
Is anchored metadata on a range inside a content control planned? If not, is there a supported place to keep structured data per content control, such as `w:dataBinding`, a custom XML part keyed by the control id, or tag conventions?
