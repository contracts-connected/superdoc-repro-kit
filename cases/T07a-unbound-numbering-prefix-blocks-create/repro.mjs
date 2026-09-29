// T07a - create.contentControl throws INVALID_INPUT "Cannot read content-control identities from malformed package XML"
// on a package whose word/numbering.xml uses w15:/w16cid: attributes without declaring those prefixes (what
// SuperDoc 1.45.x wrote on export). open, replace and save of the same package all succeed.
// Control: the same package with the prefixes declared on the <w:numbering> root.
// Second input (if present): a sanitized real template exported by 1.45.x - we try to convert the first
// {{placeholder}} of every body paragraph, as-is and after declaring the prefixes.
//
// Run: node cases/T07a-unbound-numbering-prefix-blocks-create/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { existsSync } from 'node:fs';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { header, withDoc, attempt, describe, ok, part, fixturePath, visible, verdict, range, hfStory, SuperDocClient, ROOT, CASE_ID, SDK_VERSION, sha256 } from '../../lib/harness.mjs';

const BAD = 'numbering-undeclared-prefixes';
const GOOD = 'numbering-declared-prefixes';
const REAL = 'po-header-numbering';
const hasReal = existsSync(join(ROOT, 'fixtures', 'sanitized', `${REAL}.docx`));
await header('create.contentControl on a package with undeclared namespace prefixes in numbering.xml', hasReal ? [BAD, GOOD, REAL] : [BAD, GOOD]);

const unbound = (xml) => {
  const root = xml.match(/<w:numbering\b[^>]*>/)?.[0] ?? '';
  const declared = new Set([...root.matchAll(/xmlns:(\w+)=/g)].map((m) => m[1]));
  return [...new Set([...xml.matchAll(/\b([A-Za-z]\w*):\w+=/g)].map((m) => m[1]))].filter((pfx) => pfx !== 'xmlns' && pfx !== 'xml' && !declared.has(pfx));
};
const numbering = await part(fixturePath(BAD), 'word/numbering.xml');
console.log(`${BAD}.docx numbering.xml root: ${numbering.match(/<w:numbering[^>]*>/)[0]}`);
console.log(`  prefixes used but not declared on the root: ${unbound(numbering).join(', ')}`);

// ---- 1) synthetic pair --------------------------------------------------------------------------
async function runSynthetic(fixture) {
  const { value, bytes } = await withDoc(fixture, async (doc) => ({
    // plain replace in the body: "Lorem" -> "Ipsum"
    replace: await attempt(() => doc.replace({ target: range('10000001', 0, 5), text: 'Ipsum' })),
    // inline rich-text control over the footer token "[buyer_initials]" (offsets 10..26)
    mintFooter: await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'richText', tag: 'buyer_initials', alias: 'buyer_initials', at: range('50000001', 10, 26, hfStory('rIdF1')) })),
    // inline text control over the body token "{{buyer_name}}" (offsets 7..21)
    mintBody: await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: 'buyer_name', alias: 'buyer_name', at: range('10000002', 7, 21) })),
  }), fixture);
  const docXml = await part(bytes, 'word/document.xml');
  const footer = await part(bytes, 'word/footer1.xml');
  value.replaced = visible(docXml).startsWith('Ipsum');
  value.bodySdt = /<w:tag w:val="buyer_name"\/>/.test(docXml);
  value.footerSdt = /<w:tag w:val="buyer_initials"\/>/.test(footer);
  console.log(`${fixture}: open OK | replace ${value.replaced ? 'OK' : describe(value.replace)} | body create.contentControl ${describe(value.mintBody)} (SDT in export: ${value.bodySdt}) | footer create.contentControl ${describe(value.mintFooter)} (SDT in export: ${value.footerSdt}) | save OK`);
  return value;
}
const bad = await runSynthetic(BAD);
const good = await runSynthetic(GOOD);

// ---- 2) sanitized real template -----------------------------------------------------------------
async function withPath(path, fn) {
  const client = new SuperDocClient();
  await client.connect();
  try {
    const doc = await client.open({ doc: path });
    try { return await fn(doc); } finally { await doc.close({ discard: true }).catch(() => {}); }
  } finally { await client.dispose(); }
}
// Paragraph text as the host counts it: w:t text, and <w:tab/> / <w:br/> as one character each.
function paragraphs(xml) {
  const out = [];
  for (const m of xml.matchAll(/<w:p\b[^>]*w14:paraId="([0-9A-F]{8})"[^>]*>([\s\S]*?)<\/w:p>/g)) {
    let text = '';
    for (const t of m[2].replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '').matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/g)) text += t[1] !== undefined ? t[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'") : '\t';
    out.push({ id: m[1], text });
  }
  return out;
}
async function convertFirstPlaceholders(path) {
  const xml = await part(path, 'word/document.xml');
  const targets = paragraphs(xml).map((pp) => ({ ...pp, m: pp.text.match(/\{\{[a-z0-9_]+\}\}/) })).filter((pp) => pp.m);
  return withPath(path, async (doc) => {
    const outcomes = { tried: targets.length, created: 0, invalidInput: 0, other: 0, firstError: null };
    for (const t of targets) {
      const s = t.m.index, e = s + t.m[0].length;
      const res = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: t.m[0].slice(2, -2), alias: t.m[0].slice(2, -2), at: range(t.id, s, e) }));
      if (ok(res)) outcomes.created++;
      else if (/INVALID_INPUT/.test(res.THROW ?? '') || /malformed package XML/.test(res.THROW ?? '')) { outcomes.invalidInput++; outcomes.firstError ??= res.THROW; }
      else { outcomes.other++; outcomes.firstError ??= res.THROW ?? JSON.stringify(res).slice(0, 160); }
    }
    return outcomes;
  });
}
let real = null;
if (hasReal) {
  const src = fixturePath(REAL);
  const realNumbering = await part(src, 'word/numbering.xml');
  const missing = realNumbering ? unbound(realNumbering) : [];
  console.log(`\n${REAL}.docx numbering.xml: prefixes used but not declared on the root: ${missing.join(', ') || '(none)'}`);
  // declare every missing well-known prefix on the root (only the root tag changes)
  const URIS = { w14: 'http://schemas.microsoft.com/office/word/2010/wordml', w15: 'http://schemas.microsoft.com/office/word/2012/wordml', w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid', w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex', w16: 'http://schemas.microsoft.com/office/word/2018/wordml', w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex', w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du', w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash', mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006' };
  const zip = await JSZip.loadAsync(await readFile(src));
  const fixed = realNumbering.replace(/<w:numbering\b([^>]*)>/, (m0, attrs) => `<w:numbering${attrs}${missing.filter((pfx) => URIS[pfx]).map((pfx) => ` xmlns:${pfx}="${URIS[pfx]}"`).join('')}>`);
  zip.file('word/numbering.xml', fixed);
  const dir = join(ROOT, 'out', CASE_ID); await mkdir(dir, { recursive: true });
  const declaredPath = join(dir, `${REAL}-prefixes-declared-sdk-${SDK_VERSION}.docx`);
  await writeFile(declaredPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const asIs = await convertFirstPlaceholders(src);
  const repaired = await convertFirstPlaceholders(declaredPath);
  console.log(`${REAL} as exported:          ${asIs.created}/${asIs.tried} body placeholders converted, ${asIs.invalidInput} INVALID_INPUT${asIs.firstError ? ` (first error: ${asIs.firstError})` : ''}`);
  console.log(`${REAL} prefixes declared:    ${repaired.created}/${repaired.tried} body placeholders converted, ${repaired.invalidInput} INVALID_INPUT${repaired.firstError ? ` (first other error: ${repaired.firstError})` : ''}`);
  console.log('  (one attempt per paragraph: its first {{placeholder}}; failures other than INVALID_INPUT are unrelated to this case)');
  real = { asIs, repaired };
} else {
  console.log(`\n(sanitized template fixtures/sanitized/${REAL}.docx not present - skipped)`);
}

const syntheticRepro = bad.replaced && !ok(bad.mintBody) && ok(good.mintBody) && good.bodySdt;
const realRepro = real ? real.asIs.invalidInput > 0 && real.repaired.created > real.asIs.created : false;
verdict(syntheticRepro || realRepro,
  'create.contentControl works on both packages (the host tolerates or repairs the undeclared prefix on open, as it already does for open/replace/save), or open rejects the package up front',
  `undeclared prefixes: replace ${bad.replaced ? 'OK' : 'FAIL'}, body create ${describe(bad.mintBody)}, footer create ${describe(bad.mintFooter)} | declared: body create ${describe(good.mintBody)}` +
    (real ? ` | real template: ${real.asIs.created}/${real.asIs.tried} converted as exported vs ${real.repaired.created}/${real.repaired.tried} with prefixes declared` : ''),
  { real });
