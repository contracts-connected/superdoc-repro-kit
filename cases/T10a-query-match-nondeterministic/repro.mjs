// T10a - query.match on byte-identical input, same query, N fresh client + document sessions.
// We saw `total` flip between 1 and 0 across sessions (and stay stuck for a whole session). This script counts it.
// Also: the same query with `in` omitted returns total 0 without an error, while in:{storyType:'body'} finds it.
// Fixture: one legacy bracket placeholder "[Buying Company Name]" split across two runs, the first run bold.
//
// Run: node cases/T10a-query-match-nondeterministic/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]      (N=10 by default; env N=25)
import { header, SuperDocClient, fixturePath, attempt, verdict, BODY } from '../../lib/harness.mjs';

const FIX = 'bracket-split-across-runs';
const N = Number(process.env.N ?? 10);
await header(`query.match determinism over ${N} fresh sessions, and the effect of omitting \`in\``, [FIX, 'loop-marker-in-table-cell']);

const SELECT = { type: 'text', pattern: '\\[[^\\[\\]]+\\]', mode: 'regex' };
async function session(fn, fixture = FIX) {
  const client = new SuperDocClient();
  await client.connect();
  try {
    const doc = await client.open({ doc: fixturePath(fixture) });
    try { return await fn(doc); } finally { await doc.close({ discard: true }).catch(() => {}); }
  } finally { await client.dispose(); }
}

const totals = [];
for (let i = 0; i < N; i++) {
  const r = await session((doc) => attempt(() => doc.query.match({ in: BODY, select: SELECT })));
  const repeat = r?.THROW ? null : await session(async (doc) => {
    const a = await doc.query.match({ in: BODY, select: SELECT });
    const b = await doc.query.match({ in: BODY, select: SELECT });
    return [a.total, b.total];
  });
  totals.push(r?.THROW ? `THROW ${r.THROW}` : r.total);
  console.log(`session ${String(i + 1).padStart(2)}/${N}: total=${r?.THROW ? `THROWS ${r.THROW}` : r.total}${repeat ? ` | same query twice in one session: ${repeat.join(', ')}` : ''}`);
}
const CELL = { type: 'text', pattern: '{{loop scope_items}}', mode: 'contains' };
const noIn = await session(async (doc) => ({
  withoutIn: await attempt(() => doc.query.match({ select: SELECT })),
  withBody: await attempt(() => doc.query.match({ in: BODY, select: SELECT })),
}), FIX);
const noInCell = await session(async (doc) => ({
  withoutIn: await attempt(() => doc.query.match({ select: CELL })),
  withBody: await attempt(() => doc.query.match({ in: BODY, select: CELL })),
}), 'loop-marker-in-table-cell');
const t = (x) => (x?.THROW ? `THROWS ${x.THROW}` : `total ${x.total}`);
console.log(`\nparagraph text, query.match without \`in\`: ${t(noIn.withoutIn)} | with in:body: ${t(noIn.withBody)}`);
console.log(`table-cell text, query.match without \`in\`: ${t(noInCell.withoutIn)} | with in:body: ${t(noInCell.withBody)}`);

const distinct = [...new Set(totals.map(String))];
const nondeterministic = distinct.length > 1;
const zero = (x) => !x.withoutIn?.THROW && x.withoutIn?.total === 0 && x.withBody?.total > 0;
const silentZero = zero(noIn) || zero(noInCell);
verdict(nondeterministic || silentZero,
  'the same query over the same bytes returns the same total in every session; omitting `in` either searches the body or fails with an error',
  `${N} fresh sessions returned total ${distinct.join(' / ')} (${distinct.map((d) => `${d}: ${totals.filter((x) => String(x) === d).length}`).join(', ')}); without \`in\`: paragraph ${t(noIn.withoutIn)} vs in:body ${t(noIn.withBody)}, table cell ${t(noInCell.withoutIn)} vs in:body ${t(noInCell.withBody)}`,
  { totals, nondeterministic, silentZero });
