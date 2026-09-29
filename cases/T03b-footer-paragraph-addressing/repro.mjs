// T03b - How do you address a paragraph inside a header/footer story?
// For every footer part of two inputs (a synthetic plain + table footer, and a sanitized real purchase
// order whose footers are tables inside a locked block SDT) the script asks each read API for a token
// that getText() returns, then writes to a table-cell paragraph addressed by its w14:paraId.
//
// Run: node cases/T03b-footer-paragraph-addressing/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { header, withDoc, attempt, ok, part, verdict, short, hasFn, hfStory, range, ROOT } from '../../lib/harness.mjs';

const INPUTS = [
  ['footer-plain-and-table', [['word/footer1.xml', '[buyer_initials]', '50000001'], ['word/footer2.xml', '{{contract_number}}', '61000001']]],
];
if (existsSync(join(ROOT, 'fixtures', 'sanitized', 'po-locked-footer.docx'))) {
  INPUTS.push(['po-locked-footer', [['word/footer1.xml', '{{scope_contract_number}}', null], ['word/footer2.xml', '{{scope_contract_number}}', null]]]);
}
await header('addressing text and paragraphs inside header/footer stories', INPUTS.map(([f]) => f));

const count = (r) => (r?.THROW ? `THROWS ${r.THROW}` : r?.total ?? r?.items?.length ?? r?.matches?.length ?? short(r, 80));
const n = (r) => (r?.THROW ? -1 : r?.total ?? r?.items?.length ?? r?.matches?.length ?? 0);
const paraOf = (xml, id) => (xml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${id}"[^>]*>([\\s\\S]*?)</w:p>`)) || [])[1] ?? null;
const wt = (x) => [...(x ?? '').matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');

const findings = [];
for (const [fixture, expectations] of INPUTS) {
  const src = await (await import('node:fs/promises')).readFile((await import('../../lib/harness.mjs')).fixturePath(fixture));
  console.log(`\n### ${fixture}`);
  // ---- reads (one open) ----
  const { value: reads } = await withDoc(fixture, async (doc) => {
    const parts = (await doc.headerFooters.parts.list({})).items;
    const out = [];
    for (const [partPath, token] of expectations) {
      const p = parts.find((x) => x.partPath === partPath);
      if (!p) { out.push({ partPath, missing: true }); continue; }
      const story = hfStory(p.refId);
      const sectionId = p.referencedBySections?.[0]?.sectionId;
      const text = await attempt(() => doc.getText({ in: story }));
      const match = await attempt(() => doc.query.match({ in: story, select: { type: 'text', pattern: token, mode: 'contains' } }));
      const find = hasFn(doc, 'find') ? await attempt(() => doc.find({ in: story, select: { type: 'text', pattern: token, mode: 'contains' } })) : { THROW: 'find() not present' };
      const blocks = await attempt(() => doc.blocks.list({ in: story }));
      const slotVariant = partPath === 'word/footer1.xml' && fixture === 'footer-plain-and-table' ? 'first' : 'default';
      const slot = await attempt(() => doc.query.match({ in: { kind: 'story', storyType: 'headerFooterSlot', section: { kind: 'section', sectionId }, headerFooterKind: 'footer', variant: slotVariant }, select: { type: 'text', pattern: token, mode: 'contains' } }));
      out.push({ partPath, token, refId: p.refId, sectionId, slotVariant, textHasToken: typeof text === 'string' && text.includes(token), match, find, blocks, slot });
    }
    return out;
  });
  for (const r of reads) {
    if (r.missing) { console.log(`${r.partPath}: not listed by headerFooters.parts.list`); continue; }
    const blk = r.blocks?.blocks ?? [];
    console.log(`${r.partPath} (refId ${r.refId}, token ${r.token})`);
    console.log(`   getText({ in: headerFooterPart })                     contains token: ${r.textHasToken}`);
    console.log(`   query.match({ in: headerFooterPart, text contains })   -> ${count(r.match)}`);
    console.log(`   find({ in: headerFooterPart, text contains })          -> ${count(r.find)}`);
    console.log(`   query.match({ in: headerFooterSlot ${r.sectionId}/footer/${r.slotVariant} }) -> ${count(r.slot)}`);
    console.log(`   blocks.list({ in: headerFooterPart })                  -> ${r.blocks?.THROW ? `THROWS ${r.blocks.THROW}` : `${blk.length} block(s): ${blk.map((b) => `${b.nodeType}${b.nodeId ? ' ' + b.nodeId : ''}`).join(', ')}`}`);
    if (r.textHasToken && n(r.match) === 0) findings.push(`${fixture} ${r.partPath}: query.match in headerFooterPart returns 0 for text getText returns`);
    if (r.textHasToken && n(r.find) === 0) findings.push(`${fixture} ${r.partPath}: find in headerFooterPart returns 0 for text getText returns`);
    if (r.slot?.THROW) findings.push(`${fixture} ${r.partPath}: headerFooterSlot story rejected (${r.slot.THROW})`);
    else if (r.textHasToken && n(r.slot) === 0) findings.push(`${fixture} ${r.partPath}: headerFooterSlot story returns 0`);
    const partXml = await part(src, r.partPath);
    const cellParas = partXml.includes('<w:tbl>') ? [...partXml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].flatMap((m) => [...m[0].matchAll(/w14:paraId="(\w+)"/g)].map((x) => x[1])) : [];
    if (cellParas.length && !blk.some((b) => cellParas.includes(b.nodeId))) findings.push(`${fixture} ${r.partPath}: blocks.list lists none of the ${cellParas.length} table-cell paragraphs`);
  }

  // ---- write addressed by w14:paraId (the address we rely on today) ----
  for (const [partPath, token, paraIdHint] of expectations) {
    const xml = await part(src, partPath);
    const paraId = paraIdHint ?? [...xml.matchAll(/<w:p\b[^>]*w14:paraId="(\w+)"[^>]*>([\s\S]*?)<\/w:p>/g)].find((m) => wt(m[2]).includes(token))?.[1];
    if (!paraId) continue;
    const s = wt(paraOf(xml, paraId)).indexOf(token);
    const { value: res, bytes } = await withDoc(fixture, async (doc) => {
      const p = (await doc.headerFooters.parts.list({})).items.find((x) => x.partPath === partPath);
      return attempt(() => doc.replace({ target: range(paraId, s, s + token.length, hfStory(p.refId)), text: 'VALUE' }));
    }, `${fixture}-${partPath.replace(/\W/g, '_')}-by-paraId`);
    const after = wt(paraOf(await part(bytes, partPath), paraId));
    const landed = ok(res) && after === wt(paraOf(xml, paraId)).replace(token, 'VALUE');
    console.log(`   replace() addressed by blockId = w14:paraId ${paraId} in ${partPath}: ${ok(res) ? 'success:true' : short(res?.THROW ?? res, 80)} -> exported "${after}" ${landed ? '(correct)' : '(NOT as expected)'}`);
    if (!landed) findings.push(`${fixture} ${partPath}: replace() addressed by w14:paraId ${paraId} did not land`);
  }
}

console.log('\nGaps:\n' + (findings.length ? findings.map((f) => `  - ${f}`).join('\n') : '  (none)'));
verdict(findings.length > 0,
  'text that getText() returns for a header/footer story is also found by query.match / find scoped to that story, the documented headerFooterSlot story shape is accepted, blocks.list enumerates table-cell paragraphs, and a documented paragraph address works for writes',
  findings.length ? `${findings.length} gap(s): ${findings.join('; ')}` : 'every read API finds the token; paraId addressing works',
  { gaps: findings });
