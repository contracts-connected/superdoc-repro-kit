// T06a - create.image { at: inParagraph } in a footer paragraph that also contains a PAGE complex field
// (w:fldChar) fails with TARGET_NOT_FOUND "image-paragraph-not-found", while dryRun (and capabilities.resolve
// where present) report success for the very same input. Control: the PAGE field in a sibling paragraph works.
//
// Run: node cases/T06a-create-image-fails-beside-page-field/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, findSdt, visible, verdict, short, hasFn, IMAGE_SRC, paragraph } from '../../lib/harness.mjs';

const SAME = 'footer-field-same-paragraph';
const SIBLING = 'footer-field-sibling-paragraph';
await header('create.image in a footer paragraph that contains a PAGE complex field', [SAME, SIBLING]);

// SuperDoc's suggested recipe, with capabilities.resolve and dryRun asked for the same input right before the real call.
async function recipe(doc, tag) {
  const pick = async () => (await doc.contentControls.selectByTag({ tag })).items.find((i) => i.target.story?.storyType === 'headerFooterPart');
  let control = await pick();
  const replaced = await attempt(() => doc.contentControls.replaceContent({ target: control.target, content: ' ', format: 'text' }));
  control = await pick();
  const { blockId, offset } = control.selectionTarget.end;
  const story = control.target.story;
  const input = { src: IMAGE_SRC, alt: 'Signer initials', size: { width: 72, height: 28 }, in: story, at: { kind: 'inParagraph', target: paragraph(blockId, story), offset } };
  const resolve = hasFn(doc, 'capabilities.resolve') ? await attempt(() => doc.capabilities.resolve({ operationId: 'create.image', input })) : 'n/a (no capabilities.resolve in this SDK)';
  const dryRun = await attempt(() => doc.create.image({ ...input, dryRun: true }));
  const image = await attempt(() => doc.create.image(input));
  return { replaced, input, resolve, dryRun, image };
}

async function run(fixture) {
  const { value, bytes } = await withDoc(fixture, (doc) => recipe(doc, 'buyer_initials'), fixture);
  const xml = await part(bytes, 'word/footer1.xml');
  const inside = /<w:drawing>/.test(findSdt(xml, 'buyer_initials')?.content ?? '');
  console.log(fixture);
  console.log(`  exported footer text: ${visible(xml)}`);
  console.log(`  replaceContent(' '):   ${describe(value.replaced)}`);
  console.log(`  target: paragraph ${value.input.at.target.nodeId} offset ${value.input.at.offset}`);
  console.log(`  capabilities.resolve:  ${typeof value.resolve === 'string' ? value.resolve : short(value.resolve?.available ?? value.resolve, 200)}`);
  console.log(`  create.image dryRun:   ${describe(value.dryRun)}`);
  console.log(`  create.image (real):   ${describe(value.image)}`);
  console.log(`  drawing inside buyer_initials after export: ${inside}`);
  return { ok: ok(value.image) && inside, dryOk: ok(value.dryRun), image: value.image, resolve: value.resolve };
}

const same = await run(SAME);
console.log();
const sib = await run(SIBLING);
verdict(!same.ok && sib.ok,
  'create.image succeeds in both layouts; if the field paragraph is unsupported, dryRun and capabilities.resolve say so before the real call',
  `field in same paragraph: ${describe(same.image)} (dryRun ${same.dryOk ? 'success' : 'failed'}) | field in sibling paragraph: ${sib.ok ? 'image inside SDT' : describe(sib.image)}`,
  { dryRunAgreesWithRealCall: same.dryOk === same.ok });
