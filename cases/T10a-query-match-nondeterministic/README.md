# T10a — `query.match` determinism across sessions, and the effect of omitting `in`

**Ticket:** T10 · Read paths · **Type:** Bug (not reproduced on current versions; kept as a regression check)

## Why we do this
We locate legacy placeholders (`[Buying Company Name]`, `{{loop scope_items}}`) before converting them. On earlier SDK versions, the same `query.match` on byte-identical input returned `total: 1` in some fresh sessions and `total: 0` in others, and stayed stuck within a session. Omitting `in` returned 0 for text inside a table cell. Because of that we removed `query.match` from every production path.

## What the script does
1. `N` fresh `SuperDocClient` + `open` sessions (`N=10` by default, `N=25 node … ` to raise it), each running the same regex query with `in: { kind: 'story', storyType: 'body' }` over a bracket placeholder split across two runs. It also repeats the query twice inside one session.
2. The same query without `in`, on that paragraph and on a `{{loop scope_items}}` marker inside a table cell.

```bash
node cases/T10a-query-match-nondeterministic/repro.mjs 2.15.0
N=25 node cases/T10a-query-match-nondeterministic/repro.mjs 2.15.0
```

## Expected
The same total in every session. Omitting `in` either searches everything or fails with an error; it never silently returns 0.

## Actual
Deterministic on the versions we ran: 10 of 10 sessions returned `total: 1`, and omitting `in` finds both the paragraph and the table-cell text. The script prints `FIXED?`. See [RESULTS.md](../../RESULTS.md) for every version.

## Done when
It already passes. The case stays as a regression check, and we will re-enable `query.match` once it holds on the version we ship.

## Our workaround today
We locate placeholders from the part XML and address them by `w14:paraId` plus offsets.

## Question for SuperDoc
Was there a known session-level cache behind the earlier non-determinism, fixed at a specific version? That tells us from which version `query.match` is safe on a production path.
