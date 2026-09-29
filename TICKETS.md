# Tickets

Ten root causes behind 25 runnable cases. Status per SDK version comes from [RESULTS.md](RESULTS.md)
(linux-x64; windows-x64 agrees wherever both ran). **BUG** = reproduces, **fixed** = the expected
behaviour happens, **n/a** = not a SuperDoc defect, kept as a regression check.

| Ticket | Root cause | Type | 2.13.0 | 2.15.0 | 2.16.0-next.17 | Blocks us today? |
| --- | --- | --- | --- | --- | --- | --- |
| [T01](#t01) | `create.contentControl` over text that shares a run with `<w:br/>` / `<w:tab/>` | Bug → Missing capability | BUG (silent loss) | 3 fixed, 1 BUG | 3 fixed, 1 BUG | **Yes**: our most common legacy template cannot be converted |
| [T02](#t02) | `replace()` target ranges and coordinate spaces | Bug (silent loss) | BUG | 1 fixed, 2 BUG | 1 fixed, 2 BUG | **Yes**: guarded workaround in place |
| [T03](#t03) | Header/footer story support for content controls and reads | Missing capability | BUG | BUG | BUG | **Yes**: footer initials/signatures stay as text |
| [T04](#t04) | Run properties lost when minting / filling a control | Bug | BUG | fixed | fixed | No, once we move to 2.15 |
| [T05](#t05) | Protection, permission ranges and locks | Bug + contract | BUG | 1 fixed, 2 BUG | 1 fixed, 2 BUG | Partly: our own gate enforces signing |
| [T06](#t06) | Images and offsets inside footer content controls | Bug | BUG | BUG | BUG | **Yes**: blocks the footer initials recipe |
| [T07](#t07) | Package tolerance (undeclared prefixes) | Bug | BUG | BUG | BUG | No: we repair the input before open |
| [T08](#t08) | Content-control values and identity (date text, paragraph ids) | Bug | BUG | BUG | BUG | Partly: date controls ship the raw token |
| [T09](#t09) | Tracked-change identity on mutations | Missing capability | BUG | BUG | BUG | Partly: redline attribution |
| [T10](#t10) | Read paths (determinism, text boxes, metadata in controls) | Contract / Missing capability | 2 fixed, 1 BUG | 2 fixed, 1 BUG | 2 fixed, 1 BUG | No |

Priorities from our side: **T01d, T03a, T06a/T06b, T02b, T05a**, in that order. The first three are
what our customers see; T02b and T05a are silent correctness risks we currently guard in our own code.

---

## T01

**`create.contentControl` over text that shares a run with `<w:br/>` / `<w:tab/>`**

On 2.13.0 the call returns `success:true` with a node id while deleting the break, the tab and the
neighbouring tokens (T01a, T01b), and in a footer it refuses a token that follows a tab (T01c). On
2.15.0 the silent loss is gone: the body calls refuse cleanly and the footer mint works and keeps the
formatting. What remains is the capability: a real paragraph with 110 placeholders separated by
`<w:br/>` converts **0 of 110** on every version (T01d).

| Case | 2.13.0 | 2.15.0 | 2.16.0-next.17 |
| --- | --- | --- | --- |
| [T01a](cases/T01a-create-after-break-deletes-siblings/README.md) create next to `<w:br/>` | BUG: loses `<w:br/>` and `{{f2}}` | fixed (clean refusal) | fixed (clean refusal) |
| [T01b](cases/T01b-create-before-tab-deletes-tab/README.md) create next to `<w:tab/>` | BUG: loses the tab and `{{f2}}` | fixed (clean refusal) | fixed (clean refusal) |
| [T01c](cases/T01c-footer-mint-after-tab-refused/README.md) footer token after a tab | BUG: `CAPABILITY_UNAVAILABLE` | fixed | fixed |
| [T01d](cases/T01d-real-paragraph-many-codes/README.md) real paragraph, 110 tokens | BUG: 1 created, 109 tokens lost | BUG: 0/110 convertible | BUG: 0/110 convertible |

**What we need:** minting a control over a token that shares a run with `<w:br/>` (and `<w:tab/>`),
without deleting its siblings. **Workaround today:** delete the whole span, reinsert the tokens
without breaks, mint, restore each break with `insert({ type: 'html', value: '<br>' })`. It works but
costs ~5 extra mutations per token and is fragile.

## T02

**`replace()` target ranges and coordinate spaces**

| Case | 2.13.0 | 2.15.0 | 2.16.0-next.17 |
| --- | --- | --- | --- |
| [T02a](cases/T02a-replace-then-create-loses-siblings/README.md) create on the range `replace()` reported | BUG: sibling text lost | fixed (refuses) | fixed (refuses) |
| [T02b](cases/T02b-footer-replace-overwrites-neighbours/README.md) footer `replace()` around `w:del` / `w:moveFrom` / `mc:AlternateContent` | BUG | BUG | BUG |
| [T02c](cases/T02c-tab-and-newline-from-value/README.md) a line break inside a value | BUG | BUG | BUG |

T02b: `replace()` works in the `getText` space; the `projectHtml` sourceMap reports
`coordinateSpace: "tracked"` and its ranges land 8 characters late after a `w:del` and 6 late after a
`w:moveFrom`. `replace()` accepts both with `success:true` and overwrites the neighbouring text.
**What we need:** a documented conversion between the two spaces, and an expected-text precondition on
`replace()` so a shifted range fails instead of overwriting. T02c: tabs from a value already become
`<w:tab/>`; `\n` is refused (`text-payload-unsupported-control-char`). **What we need:** a way to write a
line break from a value (multi-line addresses).

## T03

**Header/footer story support**

| Case | 2.13.0 | 2.15.0 | 2.16.0-next.17 |
| --- | --- | --- | --- |
| [T03a](cases/T03a-footer-cc-mutations-refused/README.md) mutations on a footer control | 10 of 12 refused | 6 of 12 refused, `resolve` says supported | same as 2.15.0 |
| [T03b](cases/T03b-footer-paragraph-addressing/README.md) addressing footer paragraphs | BUG | BUG | BUG |

Progress since 2.13.0: `patch`, `setLockMode`, `text.setValue` and `unwrap` now work in footers.
Still refused: `clearContent`, `appendContent`, `prependContent`, `insertAfter`, `delete`,
`text.clearValue`, while `capabilities.resolve` answers `supported`. `find()` returns 0 in every
footer story for text `getText` returns, and `blocks.list` does not enumerate paragraphs inside a
footer table. **What we need:** parity with body controls, a truthful `capabilities.resolve` meanwhile,
and confirmation that `blockId === w14:paraId` is a stable address for header/footer paragraphs.

## T04

**Run properties lost when minting / filling a control.** Fixed in 2.15.0 in both the footer
([T04a](cases/T04a-footer-mint-drops-run-properties/README.md)) and the body
([T04b](cases/T04b-body-create-setvalue-drops-run-properties/README.md)). Kept as regression checks.
Open question: are `<w:rStyle>` character-style references preserved too, or only direct formatting?

## T05

**Protection, permission ranges and locks**

| Case | 2.13.0 | 2.15.0 | 2.16.0-next.17 |
| --- | --- | --- | --- |
| [T05a](cases/T05a-readonly-does-not-gate-content-control-writes/README.md) `readOnly` vs content-control writes | BUG: 7/7 writes succeed outside the range | BUG | BUG |
| [T05b](cases/T05b-permission-ranges-not-exported/README.md) ranges after export + reopen | fixed | fixed | fixed |
| [T05c](cases/T05c-contentlocked-block-not-enforced/README.md) `contentLocked` block contents | writes succeed | writes succeed | writes succeed |

T05a: raw text writes outside the granted range are refused, but `text.setValue`, `date.setValue`,
`checkbox.setState`, `clearContent`, `create.image`, `create.contentControl` and `unwrap` all succeed
anywhere; a range cannot address a whole control. T05b works on the SDK; we saw ranges dropped by the
browser package (`superdoc@2.16.0`) export and ask you to confirm parity. T05c is a contract question:
we currently rely on writes inside a Google-Docs `contentLocked` wrapper succeeding.

## T06

**Images and offsets inside footer content controls** (your initials recipe: `selectByTag` →
`replaceContent(' ')` → re-resolve → `create.image` `inParagraph` at `selectionTarget.end`)

| Case | 2.13.0 | 2.15.0 | 2.16.0-next.17 |
| --- | --- | --- | --- |
| [T06a](cases/T06a-create-image-fails-beside-page-field/README.md) PAGE field in the same paragraph | BUG | BUG | BUG |
| [T06b](cases/T06b-second-image-lands-outside-its-sdt/README.md) two controls, second image misplaced | BUG | BUG | BUG |
| [T06c](cases/T06c-images-list-ignores-in-and-size-unit/README.md) `size.unit`, `images.list({ in })` | BUG | BUG | BUG |

**What we need:** `create.image` working in a paragraph with a complex field (or `dryRun` /
`capabilities.resolve` reporting the failure), and a target that is the content control itself
(`at: { kind: 'contentControl', target, position: 'end' }`) so callers never compute offsets.

## T07

**Package tolerance.** [T07a](cases/T07a-unbound-numbering-prefix-blocks-create/README.md): a
`numbering.xml` with `w15:` / `w16cid:` attributes and undeclared prefixes (as SuperDoc 1.45.x
exported it) opens, replaces and saves, but every `create.contentControl` throws `INVALID_INPUT Cannot
read content-control identities from malformed package XML`. Every template exported by 1.45.x carries
this. **What we need:** tolerate or repair on open, or reject at open time.
[T07b](cases/T07b-final-sectpr-loses-header-footer-references/README.md) is **not** a SuperDoc defect
(our earlier report compared against our own 1.x output); kept only as a regression check.

## T08

**Content-control values and identity**

- [T08a](cases/T08a-date-setvalue-writes-no-visible-text/README.md): `date.setValue` writes
  `w:fullDate` but no visible text, so a date control shows its original placeholder. Every version.
- [T08b](cases/T08b-paragraph-id-changes-after-block-mint-and-clear/README.md): a paragraph's id
  changes on a block mint (`00000002` → `00000004`) and again on `clearContent` (→ `00000005`); no
  receipt carries the new id, and `create.image` at the old id throws `TARGET_NOT_FOUND`. On 2.15.0
  `capabilities` now report `create.image` correctly (an older finding of ours, fixed).

## T09

**Tracked-change identity.** [T09a](cases/T09a-tracked-change-identity-and-fusion/README.md): an
`author` / `date` passed on a tracked `insert` / `replace` is silently ignored (the change goes to the
user the document was opened with), and two adjacent same-author inserts fuse into one change. **What
we need:** a per-mutation identity override, and the schema rejecting unknown fields instead of
ignoring them.

## T10

**Read paths**

- [T10a](cases/T10a-query-match-nondeterministic/README.md): `query.match` is deterministic on every
  version we tested now (10/10 sessions); our older non-determinism report was on 2.11.0. Regression check.
- [T10b](cases/T10b-textbox-content-unreachable/README.md): text inside a text box is reachable through a
  `{ storyType: 'textbox', textboxId }` story found by `query.match`; `getText` and `blocks.list` still skip
  it. Question: is that story supported, and how do we enumerate text boxes?
- [T10c](cases/T10c-metadata-attach-inside-content-control/README.md): `metadata.attach` refuses a range
  inside a content control (`text-range-in-sdt`). Missing capability.
