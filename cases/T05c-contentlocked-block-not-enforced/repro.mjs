// T05c - a first-page footer wrapped in a block SDT with <w:lock w:val="contentLocked"/> (tag goog_rdk_0, the
// wrapper Google Docs writes on export) is still editable through the Document API: create.contentControl over a
// raw token inside it, and replaceContent on an inline SDT inside it, both succeed and change the exported XML.
// Contract question: is contentLocked meant to be enforced on the block's contents?
//
// Run: node cases/T05c-contentlocked-block-not-enforced/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { readFile } from 'node:fs/promises';
import { header, withDoc, attempt, describe, ok, part, findSdt, fixturePath, visible, verdict, hfStory, range } from '../../lib/harness.mjs';

const FX = 'first-page-footer-content-locked';
await header('contentLocked block SDT in a first-page footer is not enforced', [FX]);

const lockOf = (s) => (s?.pr.match(/<w:lock w:val="(\w+)"/) || [])[1] ?? '(none)';
const inXml = await part(await readFile(fixturePath(FX)), 'word/footer1.xml');
console.log(`input footer1.xml (first-page): lock=${lockOf(findSdt(inXml, 'goog_rdk_0'))} text: ${visible(inXml)}`);

const { value, bytes } = await withDoc(FX, async (doc) => {
  const parts = (await doc.headerFooters.parts.list({})).items;
  const refId = parts.find((x) => x.partPath === 'word/footer1.xml').refId;
  const story = hfStory(refId);
  const text = await doc.getText({ in: story });
  const token = '[seller_initials]';
  // footer1 paragraph 50000001 reads "Seller initials: [seller_initials]"
  const start = 'Seller initials: '.length;
  const lockCtl = (await doc.contentControls.selectByTag({ tag: 'goog_rdk_0' })).items[0];
  const buyer = (await doc.contentControls.selectByTag({ tag: 'buyer_initials' })).items.find((i) => i.target.story?.storyType === 'headerFooterPart');
  return {
    storyText: text,
    lockReported: lockCtl?.lockMode ?? lockCtl?.properties?.lockMode ?? '(not reported)',
    // (a) mint an inline rich-text SDT over "[seller_initials]" inside the locked block
    mint: await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'richText', tag: 'seller_initials', alias: 'seller_initials', at: range('50000001', start, start + token.length, story) })),
    // (b) replace the content of the inline SDT that sits inside the locked block
    replaceInner: await attempt(() => doc.contentControls.replaceContent({ target: buyer.target, content: 'XY', format: 'text' })),
  };
}, 'edits-inside-locked-block');
const outXml = await part(bytes, 'word/footer1.xml');
const minted = !!findSdt(outXml, 'seller_initials');
const replaced = visible(findSdt(outXml, 'buyer_initials')?.content ?? '') === 'XY';
console.log(`lockMode reported by selectByTag(goog_rdk_0): ${value.lockReported}`);
console.log(`(a) create.contentControl inside the locked block:        ${describe(value.mint)} -> new SDT in exported XML: ${minted}`);
console.log(`(b) replaceContent on the inline SDT inside the locked block: ${describe(value.replaceInner)} -> content now "XY": ${replaced}`);
console.log(`exported footer1.xml: lock=${lockOf(findSdt(outXml, 'goog_rdk_0'))} text: ${visible(outXml)}`);
const aDone = ok(value.mint) && minted;
const bDone = ok(value.replaceInner) && replaced;
verdict(aDone || bDone,
  'edits inside a contentLocked block SDT are refused with a lock error, or the documentation states that the Document API does not enforce contentLocked',
  `(a) create.contentControl ${aDone ? 'SUCCEEDED' : 'refused'} | (b) replaceContent ${bDone ? 'SUCCEEDED' : 'refused'} inside a contentLocked block`,
  { mintInsideLocked: aDone, replaceInsideLocked: bDone });
