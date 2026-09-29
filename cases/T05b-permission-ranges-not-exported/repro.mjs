// T05b - readOnly protection + one permission range, then save and reopen.
// The permission range does not survive: no w:permStart / w:permEnd in the exported package, and
// permissionRanges.list() on the reopened document returns total 0. The script also reports where (if anywhere)
// the editing restriction itself is written (w:documentProtection) and what protection.get() says after reopening.
//
// Run: node cases/T05b-permission-ranges-not-exported/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import JSZip from 'jszip';
import { header, withDoc, reopen, attempt, describe, ok, verdict, short, hasFn, range } from '../../lib/harness.mjs';

const FX = 'signing-block';
await header('permission ranges and the editing restriction after export -> reopen', [FX]);

const SIG = 'buyer signature placeholder';
const { value, bytes } = await withDoc(FX, async (doc) => {
  if (!hasFn(doc, 'protection.setEditingRestriction') || !hasFn(doc, 'permissionRanges.create')) return { unsupported: true };
  const minted = await attempt(() => doc.create.contentControl({ kind: 'block', controlType: 'richText', at: range('00000002', 0, SIG.length), tag: 'buyer_signature', alias: 'Buyer Signature' }));
  const { blocks } = await doc.blocks.list({});
  const inner = blocks.find((b) => b.nodeType === 'paragraph' && b.textPreview?.includes(SIG))?.nodeId;
  const protection = await attempt(() => doc.protection.setEditingRestriction({ mode: 'readOnly' }));
  const created = await attempt(() => doc.permissionRanges.create({ target: range(inner, 0, SIG.length), principal: { kind: 'everyone' } }));
  const listBefore = await attempt(() => doc.permissionRanges.list({}));
  return { minted, protection, created, listBefore };
}, 'readonly-with-range');

if (value.unsupported) verdict(false, 'protection + permission ranges exist', 'not available in this SDK', { unsupported: true });

// Where does anything protection-related land in the exported package?
const zip = await JSZip.loadAsync(bytes);
const hits = { permStart: [], permEnd: [], documentProtection: [] };
for (const name of Object.keys(zip.files).filter((n) => n.endsWith('.xml'))) {
  const x = await zip.file(name).async('string');
  if (/<w:permStart\b/.test(x)) hits.permStart.push(name);
  if (/<w:permEnd\b/.test(x)) hits.permEnd.push(name);
  const dp = x.match(/<w:documentProtection\b[^>]*\/?>/);
  if (dp) hits.documentProtection.push(`${name}: ${dp[0]}`);
}
const after = await reopen(bytes, async (doc) => ({
  protection: await attempt(() => doc.protection.get()),
  ranges: await attempt(() => doc.permissionRanges.list({})),
}));
const total = (r) => (r?.THROW ? r.THROW : r?.total ?? r?.items?.length);

console.log(`protection.setEditingRestriction(readOnly): ${describe(value.protection)}`);
console.log(`permissionRanges.create (buyer signature text, everyone): ${describe(value.created)} ${short(value.created?.range ?? '', 160)}`);
console.log(`permissionRanges.list() before saving: total ${total(value.listBefore)}`);
console.log('\nexported package:');
console.log(`  w:permStart in:          ${hits.permStart.join(', ') || '(none)'}`);
console.log(`  w:permEnd in:            ${hits.permEnd.join(', ') || '(none)'}`);
console.log(`  w:documentProtection in: ${hits.documentProtection.join(' | ') || '(none)'}`);
console.log('\nafter reopening the exported file:');
console.log(`  protection.get():          ${short(after.protection?.editingRestriction ?? after.protection, 200)}`);
console.log(`  permissionRanges.list():   total ${total(after.ranges)} ${short(after.ranges?.items ?? '', 160)}`);

const rangeLost = ok(value.created) && Number(total(value.listBefore)) > 0 && hits.permStart.length === 0 && Number(total(after.ranges)) === 0;
const restrictionLost = after.protection?.editingRestriction?.mode !== 'readOnly';
verdict(rangeLost || restrictionLost,
  'the permission range is written as w:permStart / w:permEnd and permissionRanges.list() finds it after reopening; the readOnly restriction is written as w:documentProtection and protection.get() reports it after reopening',
  `range: ${hits.permStart.length ? 'exported' : 'NOT exported'}, list() after reopen total=${total(after.ranges)} | restriction: w:documentProtection ${hits.documentProtection.length ? 'exported' : 'NOT exported'}, protection.get() after reopen mode=${after.protection?.editingRestriction?.mode ?? short(after.protection, 80)}`,
  { rangeLost, restrictionLost, documentProtection: hits.documentProtection });
