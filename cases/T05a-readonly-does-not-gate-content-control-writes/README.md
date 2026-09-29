# T05a — `readOnly` protection refuses text edits outside a permission range, but not content-control writes

**Ticket:** T05 — protection, permission ranges and locks · **Type:** Bug

## Why we do this
During signing, a document goes out with a read-only editing restriction and one permission range per signer, covering that signer's own fields. A signer must be able to fill their own fields and nothing else. In particular, they must not be able to sign the other party's signature field or change the contract text.

## What the script does
1. Opens a six-paragraph document and mints five content controls through `create.contentControl`:
   - a block rich-text control for the buyer signature;
   - inline text, date and checkbox controls;
   - a block rich-text control for the seller signature.
2. Turns on the restriction with `protection.setEditingRestriction({ mode: 'readOnly' })`.
3. Grants one range with `permissionRanges.create`, principal `everyone`, over the text of the buyer signature.
4. **Control:** tries two raw text writes outside the range, a targetless `insert` and a `replace` on the body paragraph. Both are refused, which proves the restriction is active.
5. Tries seven content-control writes on controls outside the range and checks each one in the exported `word/document.xml`:
   - `contentControls.text.setValue`, `contentControls.date.setValue` and `contentControls.checkbox.setState`;
   - `contentControls.clearContent`, then `create.image` in the seller signature's paragraph;
   - `create.contentControl` over body text;
   - `contentControls.unwrap`, run in its own session so it does not erase the evidence of the other writes.
6. Tries `permissionRanges.create` with a whole-control target: `nodeEdge` before and after the SDT.

## Expected
With `readOnly` on and one permitted range, every write outside that range is refused, content-control writes included. A permission range can cover a whole content control.

## Actual
Raw text writes are refused:

- `document protection does not permit untargeted writes.`
- `document protection permits edits only within an everyone permission range.`

All 7 content-control writes return `success:true`, and each one is present in the exported file. `protection.get()` still reports `readOnly`, `enforced: true` and `runtimeEnforced: true` afterwards.

A `nodeEdge` range around an SDT is refused with `selection points must both be text-kind for text replace.`

See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
The script prints `FIXED?`. That means every content-control write outside the range is refused and none of them reaches the exported file.

## Our workaround today
Our application checks every write against the signer's own fields before calling the Document API. The host lock is only defence in depth. That covers our own calls, but not anything else that reaches the document.

## Question for SuperDoc
- Should an enforced editing restriction gate content-control writes, including `create.image`, `clearContent`, the `*.setValue` family, `create.contentControl` and `unwrap`?
- Can a permission range address a whole content control?
