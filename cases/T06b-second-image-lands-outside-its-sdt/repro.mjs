// T06b - two inline rich-text SDTs in one footer paragraph. SuperDoc's initials recipe applied seller first,
// then buyer: the buyer create.image returns success:true, but the exported <w:drawing> sits just BEFORE the
// buyer <w:sdtContent>. Control: buyer first, then seller -> both drawings land inside their own SDT.
// Hypothesis: selectByTag offsets count an existing inline image as zero characters, create.image counts it as one.
//
// Run: node cases/T06b-second-image-lands-outside-its-sdt/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, findSdt, visible, verdict, IMAGE_SRC, paragraph } from '../../lib/harness.mjs';

const FX = 'footer-two-initials-sdts';
await header('second image of a footer paragraph lands outside its content control', [FX]);

// SuperDoc's suggested recipe: selectByTag -> replaceContent(' ') -> re-resolve -> create.image inParagraph at
// selectionTarget.end. size.unit is omitted because the SDK schema rejects it (see T06c).
async function recipe(doc, tag) {
  const pick = async () => (await doc.contentControls.selectByTag({ tag })).items.find((i) => i.target.story?.storyType === 'headerFooterPart');
  let control = await pick();
  const replaced = await attempt(() => doc.contentControls.replaceContent({ target: control.target, content: ' ', format: 'text' }));
  control = await pick();
  const { blockId, offset } = control.selectionTarget.end;
  const story = control.target.story;
  const input = { src: IMAGE_SRC, alt: 'Signer initials', size: { width: 72, height: 28 }, in: story, at: { kind: 'inParagraph', target: paragraph(blockId, story), offset } };
  const image = await attempt(() => doc.create.image(input));
  return { tag, replaced, selectionTarget: control.selectionTarget, input, image };
}

// Walk the first footer paragraph and report where each drawing is, counting an image as 0 and as 1 character.
function layout(xml) {
  const para = xml.match(/<w:p [^>]*w14:paraId="50000001"[^>]*>([\s\S]*?)<\/w:p>/)[1];
  let pos0 = 0, pos1 = 0, inside = null; const out = [];
  for (const m of para.matchAll(/<w:tag w:val="(\w+)"\/>|<w:sdtContent>|<\/w:sdtContent>|<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:drawing>/g)) {
    if (m[1]) inside = m[1];
    else if (m[0] === '<w:sdtContent>') out.push(`  [${inside} content starts at ${pos0} (images=0) / ${pos1} (images=1)]`);
    else if (m[0] === '</w:sdtContent>') { out.push(`  [${inside} content ends   at ${pos0} (images=0) / ${pos1} (images=1)]`); inside = null; }
    else if (m[0] === '<w:drawing>') { out.push(`  DRAWING at offset ${pos0} (images=0) / ${pos1} (images=1) -> ${inside ? 'inside ' + inside : 'OUTSIDE any SDT'}`); pos1 += 1; }
    else if (m[0] === '<w:tab/>') { pos0 += 1; pos1 += 1; }
    else { pos0 += m[2].length; pos1 += m[2].length; }
  }
  return out.join('\n');
}

async function run(order) {
  const { value, bytes } = await withDoc(FX, async (doc) => { const steps = []; for (const tag of order) steps.push(await recipe(doc, tag)); return steps; }, order.join('-then-'));
  const xml = await part(bytes, 'word/footer1.xml');
  const where = Object.fromEntries(order.map((t) => [t, /<w:drawing>/.test(findSdt(xml, t)?.content ?? '')]));
  console.log(`order: ${order.join(' -> ')}`);
  for (const s of value) console.log(`  ${s.tag}: replaceContent ${describe(s.replaced)}; selectByTag selectionTarget = start ${s.selectionTarget.start.offset} / end ${s.selectionTarget.end.offset} (block ${s.selectionTarget.end.blockId}); create.image at offset ${s.input.at.offset} -> ${describe(s.image)}`);
  console.log(`  exported paragraph: ${visible(xml.match(/<w:p [^>]*w14:paraId="50000001"[^>]*>[\s\S]*?<\/w:p>/)[0])}`);
  console.log(layout(xml));
  console.log(`  drawing inside its own tagged sdtContent: ${JSON.stringify(where)}`);
  return { where, allOk: value.every((s) => ok(s.image)) };
}

const bad = await run(['seller_initials', 'buyer_initials']);
console.log();
const good = await run(['buyer_initials', 'seller_initials']);

verdict(bad.allOk && !bad.where.buyer_initials,
  'seller then buyer: each create.image success:true AND each <w:drawing> inside its own tagged <w:sdtContent> (as with buyer then seller)',
  `seller then buyer: create.image ${bad.allOk ? 'success:true x2' : 'failed'}, drawing inside ${JSON.stringify(bad.where)} | buyer then seller: inside ${JSON.stringify(good.where)}`,
  { sellerThenBuyer: bad.where, buyerThenSeller: good.where });
