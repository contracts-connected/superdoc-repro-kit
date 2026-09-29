// T05a - protection.setEditingRestriction({ mode: 'readOnly' }) + ONE permission range (principal everyone).
// Raw text writes outside the range are refused (control: the lock is on). Content-control writes are NOT
// refused anywhere: text.setValue, date.setValue, checkbox.setState, clearContent + create.image on a control
// outside the range, create.contentControl on body text, unwrap. Every "succeeded" claim is checked in the
// exported document.xml, not only in the receipt. Also: permissionRanges.create refuses a whole-SDT
// (nodeEdge) target.
//
// Run: node cases/T05a-readonly-does-not-gate-content-control-writes/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, findSdt, visible, verdict, short, hasFn, range, sel, sdtBlock, paragraph, IMAGE_SRC } from '../../lib/harness.mjs';

const FX = 'signing-block';
await header('readOnly protection + one permission range: which writes outside the range are refused?', [FX]);

const TEXT = {
  '00000001': 'Contract body text before the signature block.',
  '00000002': 'buyer signature placeholder',
  '00000003': 'buyer text placeholder',
  '00000004': 'buyer date placeholder',
  '00000005': 'buyer checkbox placeholder',
  '00000006': 'seller signature placeholder',
};
const full = (id) => range(id, 0, TEXT[id].length);

// Mint the five controls (create-first), exactly as a signing flow would.
async function mint(doc) {
  const m = async (kind, controlType, id, tag) => {
    const res = await doc.create.contentControl({ kind, controlType, at: full(id), tag, alias: tag });
    if (!ok(res)) throw new Error(`mint ${tag} failed: ${short(res)}`);
    return res.contentControl.nodeId;
  };
  return {
    buyerSig: await m('block', 'richText', '00000002', 'buyer_signature'),
    buyerText: await m('inline', 'text', '00000003', 'buyer_text'),
    buyerDate: await m('inline', 'date', '00000004', 'buyer_date'),
    buyerChk: await m('inline', 'checkbox', '00000005', 'buyer_checkbox'),
    sellerSig: await m('block', 'richText', '00000006', 'seller_signature'),
  };
}
// A block wrap gives the inner paragraph a NEW id (see T08b); find it by its text instead of the original paraId.
async function innerParagraphId(doc, text) {
  const { blocks } = await doc.blocks.list({});
  const hit = blocks.find((b) => b.nodeType === 'paragraph' && b.textPreview?.includes(text));
  if (!hit) throw new Error(`no paragraph containing "${text}"`);
  return hit.nodeId;
}

const { value, bytes } = await withDoc(FX, async (doc) => {
  if (!hasFn(doc, 'protection.setEditingRestriction') || !hasFn(doc, 'permissionRanges.create')) {
    return { unsupported: `protection.setEditingRestriction: ${hasFn(doc, 'protection.setEditingRestriction')}, permissionRanges.create: ${hasFn(doc, 'permissionRanges.create')}` };
  }
  const ids = await mint(doc);
  const protection = await attempt(() => doc.protection.setEditingRestriction({ mode: 'readOnly' }));
  // The ONE permitted range: the text of party 1's signature control (its inner paragraph).
  const buyerInner = await innerParagraphId(doc, TEXT['00000002']);
  const permitted = await attempt(() => doc.permissionRanges.create({ target: range(buyerInner, 0, TEXT['00000002'].length), principal: { kind: 'everyone' } }));

  // Controls: raw text writes outside the range - these ARE expected to be refused.
  const control = {
    'insert (targetless)': await attempt(() => doc.insert({ type: 'text', value: 'ZZZ' })),
    'replace on body paragraph': await attempt(() => doc.replace({ target: full('00000001'), text: 'CHANGED' })),
  };
  // Content-control writes on controls OUTSIDE the permitted range.
  const outside = {};
  outside['contentControls.text.setValue (buyer_text)'] = await attempt(() => doc.contentControls.text.setValue({ target: sdtBlock(ids.buyerText), value: 'OUTSIDE-TEXT' }));
  outside['contentControls.date.setValue (buyer_date)'] = await attempt(() => doc.contentControls.date.setValue({ target: sdtBlock(ids.buyerDate), value: '2026-09-22' }));
  outside['contentControls.checkbox.setState (buyer_checkbox)'] = await attempt(() => doc.contentControls.checkbox.setState({ target: sdtBlock(ids.buyerChk), checked: true }));
  // clearContent + create.image on party 2's signature control. clearContent gives the inner paragraph a new id,
  // so it is re-resolved by its position in blocks.list (same approach a signing flow has to use).
  const sellerOrdinal = (await doc.blocks.list({})).blocks.find((b) => b.nodeType === 'paragraph' && b.textPreview?.includes(TEXT['00000006']))?.ordinal;
  outside['contentControls.clearContent (seller_signature)'] = await attempt(() => doc.contentControls.clearContent({ target: sdtBlock(ids.sellerSig) }));
  const sellerInner = (await doc.blocks.list({})).blocks.find((b) => b.ordinal === sellerOrdinal)?.nodeId;
  outside['create.image in seller_signature'] = sellerInner
    ? await attempt(() => doc.create.image({ src: IMAGE_SRC, alt: 'signature', size: { width: 72, height: 28 }, at: { kind: 'inParagraph', target: paragraph(sellerInner) } }))
    : { THROW: 'could not resolve the seller control inner paragraph' };
  outside['create.contentControl on body text'] = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: 'minted_outside', at: range('00000001', 0, 8) }));
  // A whole-SDT range, which is what a signing flow would naturally grant.
  const sdtEdgeRange = await attempt(() => doc.permissionRanges.create({
    target: sel({ kind: 'nodeEdge', node: sdtBlock(ids.buyerSig), edge: 'before' }, { kind: 'nodeEdge', node: sdtBlock(ids.buyerSig), edge: 'after' }),
    principal: { kind: 'everyone' },
  }));
  const protectionAfter = await attempt(() => doc.protection.get());
  return { protection, permitted, control, outside, sdtEdgeRange, protectionAfter };
}, 'locked-writes');

// unwrap in a separate session so it does not erase the evidence above.
const unwrapRun = value.unsupported ? null : await withDoc(FX, async (doc) => {
  const ids = await mint(doc);
  await doc.protection.setEditingRestriction({ mode: 'readOnly' });
  const buyerInner = await innerParagraphId(doc, TEXT['00000002']);
  await doc.permissionRanges.create({ target: range(buyerInner, 0, TEXT['00000002'].length), principal: { kind: 'everyone' } });
  return attempt(() => doc.contentControls.unwrap({ target: sdtBlock(ids.sellerSig) }));
}, 'locked-unwrap');
if (unwrapRun) value.outside['contentControls.unwrap (seller_signature)'] = unwrapRun.value;
const unwrapXml = unwrapRun ? await part(unwrapRun.bytes, 'word/document.xml') : '';

if (value.unsupported) {
  verdict(false, 'protection + permission ranges exist', `not available in this SDK (${value.unsupported})`, { unsupported: true });
}

const xml = await part(bytes, 'word/document.xml');
const bodyPara = xml.match(/<w:p [^>]*w14:paraId="00000001"[^>]*>[\s\S]*?<\/w:p>/)?.[0] ?? '';
// Did each write reach the exported file?
const landed = {
  'insert (targetless)': xml.includes('ZZZ'),
  'replace on body paragraph': xml.includes('CHANGED'),
  'contentControls.text.setValue (buyer_text)': xml.includes('OUTSIDE-TEXT'),
  'contentControls.date.setValue (buyer_date)': /w:fullDate="2026-09-22/.test(xml),
  'contentControls.checkbox.setState (buyer_checkbox)': /<w14:checked w14:val="(1|true)"\/>/.test(findSdt(xml, 'buyer_checkbox')?.pr ?? ''),
  'contentControls.clearContent (seller_signature)': !xml.includes('seller signature placeholder'),
  'create.image in seller_signature': /<w:drawing>/.test(findSdt(xml, 'seller_signature')?.content ?? ''),
  'create.contentControl on body text': !!findSdt(bodyPara, 'minted_outside'),
  'contentControls.unwrap (seller_signature)': !findSdt(unwrapXml, 'seller_signature') && unwrapXml.includes('seller signature placeholder'),
};

console.log(`protection.setEditingRestriction(readOnly): ${describe(value.protection)} ${short(value.protection?.state?.editingRestriction ?? '', 120)}`);
console.log(`permissionRanges.create (buyer signature text, everyone): ${describe(value.permitted)} ${short(value.permitted?.range ?? value.permitted?.permissionRange ?? '', 160)}`);
const settingsXml = (await part(bytes, 'word/settings.xml')) ?? '';
console.log(`exported package carries w:documentProtection (word/settings.xml): ${/<w:documentProtection[ />]/.test(settingsXml)}`);
console.log('\ncontrols - raw text writes outside the range (expected: refused)');
for (const [k, r] of Object.entries(value.control)) console.log(`  ${k.padEnd(52)} ${describe(r).padEnd(40).slice(0, 120)}  in exported XML: ${landed[k]}`);
console.log('\ncontent-control writes outside the range (expected: refused)');
for (const [k, r] of Object.entries(value.outside)) console.log(`  ${k.padEnd(52)} ${describe(r).slice(0, 110).padEnd(40)}  in exported XML: ${landed[k]}`);
console.log(`\npermissionRanges.create with a whole-SDT (nodeEdge) target: ${describe(value.sdtEdgeRange)}`);
console.log(`protection.get() at the end: ${short(value.protectionAfter?.editingRestriction ?? value.protectionAfter, 160)}`);
console.log(`\nexported body paragraph 00000001: ${visible(bodyPara)}`);

const controlsRefused = Object.keys(value.control).every((k) => !ok(value.control[k]) && !landed[k]);
const leaked = Object.keys(value.outside).filter((k) => ok(value.outside[k]) && landed[k]);
verdict(leaked.length > 0,
  'with readOnly on and one permitted range, every write outside that range is refused - raw text writes AND content-control writes; a range can target a whole content control',
  `raw text writes outside: ${controlsRefused ? 'refused (lock is on)' : 'NOT refused'} | content-control writes outside that succeeded and reached the file: ${leaked.length}/${Object.keys(value.outside).length} (${leaked.map((k) => k.split(' (')[0].replace('contentControls.', '')).join(', ')}) | nodeEdge range: ${ok(value.sdtEdgeRange) ? 'accepted' : 'refused'}`,
  { leaked, controlsRefused, sdtEdgeRange: describe(value.sdtEdgeRange) });
