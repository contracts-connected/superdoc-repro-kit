// T04a - create.contentControl (inline rich text) over a styled token: the SDT is created, but the run
// inside the new <w:sdtContent> has no <w:rPr> - font, bold, colour and size of the wrapped text are lost.
// Three shapes, each on a fresh open: footer token sharing a run with its label, footer token in its own
// run, and the same-run shape in the body for comparison. Verdict from the exported XML.
//
// Run: node cases/T04a-footer-mint-drops-run-properties/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { readFile } from 'node:fs/promises';
import { header, withDoc, attempt, describe, ok, part, findSdt, rPrOf, visible, verdict, fixturePath, hasFn, BODY, hfStory, range } from '../../lib/harness.mjs';

const FIX = 'styled-tokens-footer-and-body';
await header('minted content control keeps the wrapped run\'s <w:rPr>?', [FIX]);

// [label, partName, paraId, token, tag, storyKind]
const SHAPES = [
  ['footer, token in the same run as its label', 'word/footer1.xml', '50000001', '[buyer_initials]', 'buyer_initials', 'footer'],
  ['footer, token in its own run', 'word/footer1.xml', '50000002', '[seller_initials]', 'seller_initials', 'footer'],
  ['body, token in the same run as its label (comparison)', 'word/document.xml', '10000001', '[body_initials]', 'body_initials', 'body'],
];
const src = await readFile(fixturePath(FIX));
const rows = [];
for (const [label, partName, paraId, token, tag, kind] of SHAPES) {
  const inXml = await part(src, partName);
  const tokenRun = [...inXml.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((m) => m[0]).find((x) => x.includes(token));
  const { value, bytes } = await withDoc(FIX, async (doc) => {
    const story = kind === 'body' ? BODY : hfStory((await doc.headerFooters.parts.list({})).items[0].refId);
    const preview = (await doc.blocks.list({ in: story })).blocks.find((b) => b.nodeId === paraId).textPreview;
    const s = preview.indexOf(token);
    const input = { kind: 'inline', controlType: 'richText', tag, alias: tag, at: range(paraId, s, s + token.length, kind === 'body' ? undefined : story) };
    const resolve = hasFn(doc, 'capabilities.resolve') ? await attempt(() => doc.capabilities.resolve({ operationId: 'create.contentControl', input })) : null;
    const result = await attempt(() => doc.create.contentControl(input));
    return { preview, s, resolve, result };
  }, tag);
  const outXml = await part(bytes, partName);
  const sdt = findSdt(outXml, tag);
  const before = rPrOf(tokenRun);
  const after = sdt ? rPrOf(sdt.content) : '(no SDT in export)';
  const kept = !!sdt && after === before;
  rows.push({ label, ok: ok(value.result), created: !!sdt, kept, before, after });
  console.log(`${label}`);
  console.log(`   paragraph ${paraId} text ${JSON.stringify(value.preview)}  selection [${value.s},${value.s + token.length})  resolve: ${value.resolve == null ? 'n/a' : JSON.stringify(value.resolve?.available ?? value.resolve)}`);
  console.log(`   create.contentControl: ${describe(value.result)}   SDT in export: ${sdt ? 'yes' : 'NO'}`);
  console.log(`   rPr of the token run before: ${before}`);
  console.log(`   rPr inside the new SDT:      ${after}   -> ${kept ? 'KEPT' : 'LOST'}`);
  if (sdt) console.log(`   exported SDT: ${sdt.xml}`);
}

const footerLost = rows.filter((x) => x.label.startsWith('footer') && x.created && !x.kept);
const footerFailed = rows.filter((x) => x.label.startsWith('footer') && !x.created);
const body = rows.find((x) => x.label.startsWith('body'));
verdict(footerLost.length > 0 || footerFailed.length > 0,
  'each mint succeeds, the SDT wraps exactly the token, and the run inside <w:sdtContent> keeps the original <w:rPr> (Arial, bold, 0B5394, sz 20)',
  `footer: ${rows.filter((x) => x.label.startsWith('footer')).map((x) => `${x.label.includes('own run') ? 'own-run' : 'same-run'} ${x.created ? (x.kept ? 'rPr kept' : 'rPr LOST') : 'not created'}`).join(', ')}; body comparison: ${body.created ? (body.kept ? 'rPr kept' : 'rPr LOST') : 'not created'}`,
  { shapes: rows.map(({ label, created, kept, before, after }) => ({ label, created, rPrKept: kept, before, after })) });
