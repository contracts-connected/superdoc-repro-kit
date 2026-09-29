# T07a — Undeclared `w15`/`w16cid` prefixes in numbering.xml make `create.contentControl` throw

**Ticket:** T07 · Package tolerance and export fidelity · **Type:** Bug

## Why we do this
Our customers' templates were exported by SuperDoc 1.45.x. That exporter copied the source `<w:numbering>` root verbatim and then wrote child elements carrying `w15:restartNumberingAfterBreak` and `w16cid:durableId`, without declaring those prefixes. Every such template opens, edits and saves fine on 2.x, but the first content-control call fails. Converting `{{placeholders}}` into content controls is the first thing we do with a template, so the whole template becomes unusable.

## What the script does
1. Builds two synthetic packages that differ only in the `<w:numbering>` root: one declares only `xmlns:w` (as 1.45.x wrote it), the other also declares `xmlns:w15` and `xmlns:w16cid`.
2. On each: a plain `replace()` in the body, then `create.contentControl` over a body token and over a footer token, then `save`. The exported XML is checked for the new SDTs.
3. If `fixtures/sanitized/po-header-numbering.docx` is present (a sanitized real template exported by 1.45.x), it tries to convert the first `{{placeholder}}` of every body paragraph, as exported and again after declaring the missing prefixes on the numbering root.

```bash
node cases/T07a-unbound-numbering-prefix-blocks-create/build.mjs   # rebuild fixtures (optional)
node cases/T07a-unbound-numbering-prefix-blocks-create/repro.mjs 2.15.0
```

## Expected
`create.contentControl` works on both packages, since the host already tolerates the prefixes for open, replace and save. Alternatively, `open` rejects the package up front instead of failing at the first content-control call.

## Actual
On the undeclared package, `replace` and `save` succeed. Both `create.contentControl` calls throw `INVALID_INPUT Cannot read content-control identities from malformed package XML.` On the declared package both succeed. See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
The script prints `FIXED?`: `create.contentControl` succeeds on the undeclared package, or `open` rejects it.

## Our workaround today
Before `client.open()`, we rewrite the `<w:numbering>` root tag to declare every well-known prefix that is used but not declared (`w15`, `w16cid`, …). Only the root tag changes. Without it, a real template converted 0 of 14 body placeholders.

## Question for SuperDoc
Can the content-control identity scanner tolerate an unbound well-known prefix, the way every other operation does? Or can the host repair it on open? If neither, could `open` fail explicitly for such a package?
