// T03a - contentControls.* mutations on a FOOTER content control.
// Only getContent and replaceContent work; every other mutation throws CAPABILITY_UNAVAILABLE
// "... outside the document body ...", while capabilities.resolve (SDK 2.14.0+) reports them as supported.
// Each operation runs on a fresh open of the fixture; the effect is checked in the exported word/footer1.xml.
//
// Run: node cases/T03a-footer-cc-mutations-refused/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, part, findSdt, visible, verdict, short, hasFn } from '../../lib/harness.mjs';

const RICH = ['footer-two-rich-sdts', 'seller_initials'];
const PLAIN = ['footer-plain-text-sdt', 'initials'];
await header('content-control mutations on a footer SDT vs capabilities.resolve', [RICH[0], PLAIN[0]]);

const lockOf = (s) => (s.pr.match(/<w:lock w:val="(\w+)"/) || [])[1] ?? null;
// op -> [[fixture, tag], input(target), check(exportedSdt, footerXml) -> did the op take effect?]
const OPS = {
  'contentControls.getContent': [RICH, (t) => ({ target: t }), null],
  'contentControls.replaceContent': [RICH, (t) => ({ target: t, content: 'AB', format: 'text' }), (s) => s && visible(s.content) === 'AB'],
  'contentControls.clearContent': [RICH, (t) => ({ target: t }), (s) => s && !visible(s.content).includes('[seller_initials]')],
  'contentControls.appendContent': [RICH, (t) => ({ target: t, content: 'Z', format: 'text' }), (s) => s && visible(s.content).endsWith('Z')],
  'contentControls.prependContent': [RICH, (t) => ({ target: t, content: 'A', format: 'text' }), (s) => s && visible(s.content).startsWith('A')],
  'contentControls.patch': [RICH, (t) => ({ target: t, alias: 'Patched' }), (s) => s && s.pr.includes('w:alias w:val="Patched"')],
  'contentControls.setLockMode': [RICH, (t) => ({ target: t, lockMode: 'unlocked' }), (s) => s && lockOf(s) == null],
  'contentControls.insertAfter': [RICH, (t) => ({ target: t, content: 'AFTER', format: 'text' }), (s, x) => visible(x).includes('AFTER')],
  'contentControls.unwrap': [RICH, (t) => ({ target: t }), (s, x) => !s && visible(x).includes('[seller_initials]')],
  'contentControls.delete': [RICH, (t) => ({ target: t }), (s) => !s],
  'contentControls.text.setValue': [PLAIN, (t) => ({ target: t, value: 'JD' }), (s) => s && visible(s.content) === 'JD'],
  'contentControls.text.clearValue': [PLAIN, (t) => ({ target: t }), (s) => s && visible(s.content) === ''],
};

const rows = [];
for (const [op, [[fixture, tag], mkInput, check]] of Object.entries(OPS)) {
  const { value, bytes } = await withDoc(fixture, async (doc) => {
    const c = (await doc.contentControls.selectByTag({ tag })).items.find((i) => i.target.story?.storyType === 'headerFooterPart');
    const input = mkInput(c.target);
    const resolve = hasFn(doc, 'capabilities.resolve') ? await attempt(() => doc.capabilities.resolve({ operationId: op, input })) : null;
    let fn = doc; for (const k of op.split('.')) fn = fn[k];
    const result = typeof fn === 'function' ? await attempt(() => fn(input)) : { THROW: `${op} is not a function in this SDK` };
    return { resolve, result };
  }, check ? op.replace(/\./g, '_') : null);
  const xml = bytes ? await part(bytes, 'word/footer1.xml') : null;
  const effect = check ? !!check(findSdt(xml, tag), xml) : null;
  const works = ok(value.result) && (check ? effect : true);
  const says = value.resolve == null ? 'n/a' : value.resolve.THROW ? 'THROWS' : value.resolve.available?.kind ?? value.resolve.kind ?? short(value.resolve, 60);
  rows.push({ op, works, says });
  console.log(`${op.padEnd(34)} resolve=${String(says).padEnd(11)} actual=${works ? 'WORKS  ' : 'REFUSED'} ${works ? '' : (value.result?.THROW ?? `no effect in exported XML (${short(value.result, 120)})`)}`);
}

const refused = rows.filter((x) => !x.works);
const lying = refused.filter((x) => x.says === 'supported');
verdict(refused.length > 0,
  'every contentControls.* mutation works on a footer SDT as on a body SDT, or capabilities.resolve reports the refused ones as unsupported',
  `${rows.length - refused.length}/${rows.length} work; ${refused.length} refused; capabilities.resolve says "supported" for ${lying.length} of the refused ops${rows[0].says === 'n/a' ? ' (no capabilities.resolve in this SDK)' : ''}`,
  { refused: refused.map((x) => x.op), resolveSupportedButRefused: lying.map((x) => x.op) });
