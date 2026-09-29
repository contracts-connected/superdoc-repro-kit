// T06c - two small image contract issues on the SDK host:
//  (a) create.image rejects size.unit ('px'), which the browser package accepts and which images.list itself returns;
//  (b) images.list({ in: <footer story> }) is not scoped to that story: an image that exists only in the BODY is listed.
//
// Run: node cases/T06c-images-list-ignores-in-and-size-unit/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, short, verdict, IMAGE_SRC, paragraph, hfStory } from '../../lib/harness.mjs';

const FX = 'footer-initials-sdt';
await header('create.image size.unit on the SDK; images.list story scoping', [FX]);

// (a) SuperDoc's initials recipe verbatim, i.e. WITH size.unit: 'px'
const a = await withDoc(FX, async (doc) => {
  const pick = async () => (await doc.contentControls.selectByTag({ tag: 'seller_initials' })).items.find((i) => i.target.story?.storyType === 'headerFooterPart');
  const c = await pick();
  await doc.contentControls.replaceContent({ target: c.target, content: ' ', format: 'text' });
  const c2 = await pick();
  const story = c2.target.story; const { blockId, offset } = c2.selectionTarget.end;
  return attempt(() => doc.create.image({ src: IMAGE_SRC, alt: 'Signer initials', size: { width: 72, height: 28, unit: 'px' }, in: story, at: { kind: 'inParagraph', target: paragraph(blockId, story), offset } }));
}, 'a-size-unit');
const aXml = await part(a.bytes, 'word/footer1.xml');
console.log(`(a) create.image with size { width: 72, height: 28, unit: 'px' } -> ${describe(a.value)}; drawing in exported footer: ${/<w:drawing>/.test(aXml)}`);

// (b) put ONE image in the body only, then list images scoped to the footer story
const b = await withDoc(FX, async (doc) => {
  const refId = (await doc.headerFooters.parts.list({})).items.find((x) => x.partPath === 'word/footer1.xml')?.refId ?? 'rIdF1';
  const FOOTER = hfStory(refId);
  const put = await attempt(() => doc.create.image({ src: IMAGE_SRC, alt: 'Body image', size: { width: 72, height: 28 }, at: { kind: 'inParagraph', target: paragraph('10000001'), offset: 0 } }));
  const all = await attempt(() => doc.images.list({}));
  const inFooter = await attempt(() => doc.images.list({ in: FOOTER }));
  return { put, all, inFooter, FOOTER };
}, 'b-images-list');
const docXml = await part(b.bytes, 'word/document.xml');
const ftrXml = await part(b.bytes, 'word/footer1.xml');
const n = (r) => (r?.THROW ? r.THROW : r?.total ?? r?.items?.length);
console.log(`(b) body image inserted: ${describe(b.value.put)}; exported: drawing in document.xml=${/<w:drawing>/.test(docXml)}, in footer1.xml=${/<w:drawing>/.test(ftrXml)}`);
console.log(`    images.list({})                              -> total ${n(b.value.all)}`);
console.log(`    images.list({ in: ${JSON.stringify(b.value.FOOTER)} }) -> total ${n(b.value.inFooter)}`);
console.log(`    first item of the footer-scoped list: ${short(b.value.inFooter?.items?.[0], 300)}`);

const aBug = !ok(a.value);
const bBug = ok(b.value.put) && !/<w:drawing>/.test(ftrXml) && (Number(n(b.value.inFooter)) || 0) > 0;
verdict(aBug || bBug,
  "(a) size.unit accepted on the SDK as in the browser package (or documented as SDK-unsupported); (b) images.list({in: footer}) returns 0 when the only image is in the body",
  `(a) ${ok(a.value) ? 'size.unit accepted' : describe(a.value)} | (b) images.list({in: footer}) total=${n(b.value.inFooter)} while the only image is in the body`,
  { sizeUnitRejected: aBug, imagesListIgnoresIn: bBug });
