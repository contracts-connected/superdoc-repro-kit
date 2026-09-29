// T10c - anchored metadata (doc.metadata.attach) on a range INSIDE an inline content control.
// Control: the same call on plain body text outside any control.
// Paragraph 00000001: plain text. Paragraph 00000002: "Buyer initials: " + inline rich-text SDT tagged buyer_initials ("AB").
//
// Run: node cases/T10c-metadata-attach-inside-content-control/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, verdict, range, hasFn, short } from '../../lib/harness.mjs';

const FIX = 'inline-sdt-in-body';
await header('metadata.attach inside a content control vs on plain text', [FIX]);

const { value, bytes } = await withDoc(FIX, async (doc) => {
  const v = { api: hasFn(doc, 'metadata.attach') };
  if (!v.api) { console.log('doc.metadata.attach is not present in this SDK'); return v; }
  v.outside = await attempt(() => doc.metadata.attach({ target: range('00000001', 0, 5), namespace: 'example-signing', payload: { field: 'outside' } }));
  const sdt = (await doc.contentControls.selectByTag({ tag: 'buyer_initials' })).items[0];
  v.sdtTarget = sdt?.selectionTarget;
  v.inside = await attempt(() => doc.metadata.attach({ target: sdt.selectionTarget, namespace: 'example-signing', payload: { field: 'buyer_initials' } }));
  v.list = hasFn(doc, 'metadata.list') ? await attempt(() => doc.metadata.list({ namespace: 'example-signing' })) : null;
  console.log(`metadata.attach on plain text (paragraph 00000001, 0..5): ${describe(v.outside)}`);
  console.log(`selectByTag(buyer_initials).selectionTarget = ${short(v.sdtTarget, 300)}`);
  console.log(`metadata.attach on that selectionTarget (inside the control): ${describe(v.inside)}`);
  console.log(`metadata.list: ${short(v.list, 300)}`);
  return v;
}, 'after-attach');

if (!value.api) verdict(true, 'doc.metadata.attach exists and accepts a range inside a content control', 'doc.metadata.attach is not present', {});
const xml = await part(bytes, 'word/document.xml');
console.log(`exported document.xml mentions the payload: outside=${xml.includes('outside') || JSON.stringify(value.list ?? {}).includes('outside')}, inside=${xml.includes('buyer_initials') && JSON.stringify(value.list ?? {}).includes('"buyer_initials"')}`);
verdict(ok(value.outside) && !ok(value.inside),
  'metadata.attach accepts a range inside a content control (to carry a field identity), as it does on plain text',
  `plain text: ${describe(value.outside)} | inside the control: ${describe(value.inside)}`,
  {});
