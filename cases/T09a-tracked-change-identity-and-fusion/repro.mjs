// T09a - tracked-change identity: a tracked mutation is always attributed to the user the document was opened with.
// There is no per-mutation author/date/id; passing one is rejected or ignored (the script shows which).
// Also: two consecutive tracked inserts by the same user, the second right where the first ended, become ONE change.
// Paragraph 00000001: "Please review the original phrase before signing." (insert point: offset 14, after "Please review ")
//
// Run: node cases/T09a-tracked-change-identity-and-fusion/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, verdict, range, short } from '../../lib/harness.mjs';

const FIX = 'review-paragraph';
const USER = { userName: 'Alice Reviewer', userEmail: 'alice@example.test' };
await header('tracked mutations: per-call author/date override and same-author fusion', [FIX]);

const summarize = (list) => (list?.items ?? []).map((i) => ({ id: i.id, type: i.type, author: i.author, authorEmail: i.authorEmail, date: i.date, excerpt: i.excerpt ?? i.text ?? i.preview }));

const { value, bytes } = await withDoc(FIX, async (doc) => {
  const v = {};
  v.list0 = summarize(await doc.trackChanges.list());
  console.log(`opened with ${JSON.stringify(USER)}; tracked changes: ${v.list0.length}`);

  // 1) try to attribute a tracked insert to someone else, with an explicit date
  v.override = await attempt(() => doc.insert({ type: 'text', value: 'BOB ', target: range('00000001', 0, 0), changeMode: 'tracked', author: 'Bob Counterparty', authorEmail: 'bob@example.test', date: '2020-01-01T00:00:00Z' }));
  console.log(`1. insert(..., changeMode:'tracked', author:'Bob Counterparty', date:'2020-01-01...') -> ${describe(v.override)}${v.override?.THROW ? '' : ` (receipt keys: ${Object.keys(v.override ?? {}).join(', ')})`}`);
  v.list1 = summarize(await doc.trackChanges.list());
  if (ok(v.override)) console.log(`   tracked changes now: ${JSON.stringify(v.list1)}`);
  const off = ok(v.override) ? 4 : 0; // "BOB " shifts the paragraph when it was accepted

  // 2) two consecutive tracked inserts by the same user, the second exactly where the first ended
  const before = new Set(v.list1.map((i) => i.id));
  v.ins1 = await attempt(() => doc.insert({ type: 'text', value: 'CAREFULLY ', target: range('00000001', off + 14, off + 14), changeMode: 'tracked' }));
  v.ins2 = await attempt(() => doc.insert({ type: 'text', value: 'AGAIN ', target: range('00000001', off + 24, off + 24), changeMode: 'tracked' }));
  console.log(`2. insert 'CAREFULLY ' at ${off + 14} -> ${describe(v.ins1)}; insert 'AGAIN ' at ${off + 24} -> ${describe(v.ins2)}`);
  v.list2 = summarize(await doc.trackChanges.list());
  v.newChanges = v.list2.filter((i) => !before.has(i.id));
  console.log(`   new tracked changes after the two inserts: ${v.newChanges.length} -> ${JSON.stringify(v.newChanges)}`);
  return v;
}, 'tracked', USER);

const xml = await part(bytes, 'word/document.xml');
const ins = [...xml.matchAll(/<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>/g)].map((m) => ({ author: m[1].match(/w:author="([^"]*)"/)?.[1], date: m[1].match(/w:date="([^"]*)"/)?.[1], text: [...m[2].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join('') }));
console.log(`exported <w:ins>: ${JSON.stringify(ins)}`);

const overrideHonoured = ins.some((i) => i.author === 'Bob Counterparty');
const overrideOutcome = value.override?.THROW ? `rejected (${value.override.THROW.slice(0, 90)})` : overrideHonoured ? 'honoured' : `ignored (change attributed to ${JSON.stringify(ins.find((i) => i.text.includes('BOB'))?.author ?? value.list1.find((i) => !value.list0.some((j) => j.id === i.id))?.author)})`;
const fused = value.newChanges.length === 1 && ok(value.ins1) && ok(value.ins2);
verdict(!overrideHonoured,
  'a tracked insert/replace/delete can carry an explicit author, authorEmail and date (as OOXML w:ins/w:del do), so a change can be re-applied on behalf of the reviewer who made it',
  `per-call author/date: ${overrideOutcome}; exported authors: ${[...new Set(ins.map((i) => i.author))].join(', ') || '(no w:ins)'}; two adjacent same-author tracked inserts -> ${value.newChanges.length} change(s)${fused ? ' (fused)' : ''}`,
  { overrideOutcome, fused, exportedIns: ins });
