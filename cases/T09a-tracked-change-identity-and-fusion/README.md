# T09a — Tracked mutations cannot carry an author or date, and adjacent inserts fuse

**Ticket:** T09 · Tracked-change identity · **Type:** Missing capability

## Why we do this
Our review flow persists each tracked change with its author and date. It sometimes has to re-apply a change on behalf of the reviewer who made it, for example when restoring a change after an undo or carrying a review round into a new document version. OOXML stores author and date on every `<w:ins>`/`<w:del>`, but the Document API only lets the session's user own a change.

## What the script does
1. Opens the document with `client.open({ doc, userName: 'Alice Reviewer', userEmail: 'alice@example.test' })`.
2. Calls a tracked `insert` that also passes `author: 'Bob Counterparty'`, `authorEmail` and `date: '2020-01-01T00:00:00Z'`. It prints whether those fields are rejected, ignored or honoured.
3. Makes two tracked inserts by the same user, the second exactly where the first ended, and counts how many tracked changes they produce.
4. Reads `w:author` and `w:date` of every `<w:ins>` in the exported `word/document.xml`.

```bash
node cases/T09a-tracked-change-identity-and-fusion/repro.mjs 2.15.0
```

## Expected
A tracked insert, replace or delete can carry an explicit author, author email and date, which end up on the exported `<w:ins>`/`<w:del>`. Merging adjacent same-author inserts is fine, but we would like it documented or controllable.

## Actual
The extra fields are accepted without error and ignored: the change is attributed to "Alice Reviewer" with the current time. The two adjacent inserts become one change ("CAREFULLY AGAIN "). See [RESULTS.md](../../RESULTS.md).

## Done when
The script prints `FIXED?`: the exported `<w:ins>` carries `w:author="Bob Counterparty"`.

## Our workaround today
We keep our own mapping between persisted change ids and SuperDoc's ids. A re-applied change is attributed to whoever has the document open, which is wrong in the review history.

## Question for SuperDoc
Is a per-mutation identity override (`author`, `authorEmail`, `date`, and ideally the `w:id`) planned for `insert`, `replace` and `delete` with `changeMode: 'tracked'`? Should unknown fields be rejected by the schema rather than silently ignored? Is adjacent-insert fusion documented, and can it be turned off per call?
