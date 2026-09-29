// T08b - paragraph identity is not stable across a block content-control mint and a clearContent:
//   1. blocks.list shows the paragraph as nodeId 00000002 (its w14:paraId)
//   2. create.contentControl kind 'block' over that paragraph -> the paragraph now inside the SDT has a NEW nodeId
//   3. contentControls.clearContent on that SDT -> the inner paragraph gets yet another nodeId
// Neither receipt returns the new paragraph id, so a caller must re-list blocks before create.image inParagraph.
// Also printed: what capabilities says about create.image vs create.image actually succeeding in that paragraph.
//
// Run: node cases/T08b-paragraph-id-changes-after-block-mint-and-clear/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, findSdt, verdict, range, sdtBlock, paragraph, short, hasFn, IMAGE_SRC } from '../../lib/harness.mjs';

const FIX = 'three-paragraphs';
const TEXT = 'buyer signature placeholder';
await header('paragraph nodeId before/after a block content-control mint and after clearContent', [FIX]);

const listParas = async (doc) => ((await doc.blocks.list({})).blocks ?? []).filter((b) => b.nodeType === 'paragraph');
const idOf = (paras, predicate) => paras.find(predicate)?.nodeId ?? null;

const { value, bytes } = await withDoc(FIX, async (doc) => {
  const v = {};
  const p0 = await listParas(doc);
  v.before = idOf(p0, (b) => (b.textPreview ?? b.text ?? '').startsWith(TEXT));
  const ordinal = p0.find((b) => b.nodeId === v.before)?.ordinal;
  console.log(`1. blocks.list: "${TEXT}" is paragraph nodeId=${v.before} (ordinal ${ordinal})`);

  v.mint = await attempt(() => doc.create.contentControl({ kind: 'block', controlType: 'richText', tag: 'buyer_signature', alias: 'Buyer Signature', at: range(v.before, 0, TEXT.length) }));
  console.log(`2. create.contentControl kind:block -> ${describe(v.mint)} receipt=${short(v.mint, 400)}`);
  const p1 = await listParas(doc);
  v.afterMint = idOf(p1, (b) => (b.textPreview ?? b.text ?? '').startsWith(TEXT));
  v.oldIdStillListed = p1.some((b) => b.nodeId === v.before);
  console.log(`   blocks.list: that paragraph is now nodeId=${v.afterMint}; original id ${v.before} still listed: ${v.oldIdStillListed}`);
  console.log(`   paragraphs now: ${p1.map((b) => `${b.ordinal}:${b.nodeId}:"${(b.textPreview ?? '').slice(0, 24)}"`).join(' | ')}`);

  const sdtId = v.mint?.contentControl?.nodeId;
  v.clear = await attempt(() => doc.contentControls.clearContent({ target: sdtBlock(sdtId) }));
  console.log(`3. contentControls.clearContent -> ${describe(v.clear)} receipt=${short(v.clear, 400)}`);
  const p2 = await listParas(doc);
  const inner = p2.find((b) => b.ordinal === p1.find((x) => x.nodeId === v.afterMint)?.ordinal) ?? null;
  v.afterClear = inner?.nodeId ?? null;
  console.log(`   blocks.list: the (now empty) paragraph at the same ordinal is nodeId=${v.afterClear}`);
  console.log(`   paragraphs now: ${p2.map((b) => `${b.ordinal}:${b.nodeId}:"${(b.textPreview ?? '').slice(0, 24)}"`).join(' | ')}`);

  // capabilities vs reality for create.image
  const input = { src: IMAGE_SRC, alt: 'Signature', size: { width: 72, height: 28 }, at: { kind: 'inParagraph', target: paragraph(v.afterClear) } };
  if (hasFn(doc, 'capabilities.get')) {
    const caps = await attempt(() => doc.capabilities.get());
    const s = JSON.stringify(caps);
    const i = s.indexOf('create.image');
    v.capsGetImage = i >= 0 ? s.slice(i, i + 160) : '(create.image not mentioned)';
    console.log(`4. capabilities.get(): ${v.capsGetImage}`);
  }
  if (hasFn(doc, 'capabilities.resolve')) {
    v.capsResolveImage = await attempt(() => doc.capabilities.resolve({ operationId: 'create.image', input }));
    console.log(`   capabilities.resolve(create.image, that paragraph): ${short(v.capsResolveImage, 200)}`);
  }
  v.imageOldId = await attempt(() => doc.create.image({ ...input, at: { kind: 'inParagraph', target: paragraph(v.before) } }));
  console.log(`5. create.image inParagraph at the ORIGINAL id ${v.before}: ${describe(v.imageOldId)}`);
  v.imageNewId = await attempt(() => doc.create.image(input));
  console.log(`   create.image inParagraph at the CURRENT id ${v.afterClear}: ${describe(v.imageNewId)}`);
  return v;
}, 'after-mint-clear-image');

const xml = await part(bytes, 'word/document.xml');
const sdt = findSdt(xml, 'buyer_signature');
const innerParaId = sdt?.content.match(/w14:paraId="([0-9A-F]+)"/)?.[1] ?? null;
const drawingInside = !!sdt && sdt.content.includes('<w:drawing>');
console.log(`exported: SDT buyer_signature inner paragraph w14:paraId=${innerParaId}; drawing inside the SDT: ${drawingInside}`);

const changed = [value.before !== value.afterMint ? `mint ${value.before}->${value.afterMint}` : null, value.afterMint !== value.afterClear ? `clearContent ${value.afterMint}->${value.afterClear}` : null].filter(Boolean);
const receiptsCarryId = [value.mint, value.clear].some((r) => JSON.stringify(r ?? {}).includes(String(value.afterClear)) || JSON.stringify(r ?? {}).includes(String(value.afterMint)));
verdict(changed.length > 0 && !receiptsCarryId,
  'the paragraph keeps its id through a block mint and clearContent, or the receipts return the new inner paragraph id',
  `paragraph id changed on: ${changed.join(', ') || 'nothing'}; receipts carry the new id: ${receiptsCarryId}; create.image at original id: ${describe(value.imageOldId)}; at current id: ${describe(value.imageNewId)}${value.capsGetImage ? `; capabilities.get: ${value.capsGetImage.slice(0, 80)}` : ''}`,
  { ids: { before: value.before, afterMint: value.afterMint, afterClear: value.afterClear, exportedInner: innerParaId }, drawingInside });
