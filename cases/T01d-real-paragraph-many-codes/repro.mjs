// T01d - the real paragraph from our field-catalog template: 110 {{placeholders}} in ONE run, separated by
// <w:br/> (mostly two breaks between tokens) with a leading break. Our import wants to turn every placeholder
// into an inline content control. Three passes, each on a fresh open:
//   A. create.contentControl on the FIRST token only
//   B. create.contentControl on the LAST token only
//   C. the import itself: every token, last to first (so earlier offsets stay valid)
// The outcome is read from the exported word/document.xml: are all 110 tokens and all breaks still there,
// and does each new SDT wrap exactly its token?
//
// Fixture: fixtures/sanitized/paragraph-116-codes.docx (sanitized real template).
// Run: node cases/T01d-real-paragraph-many-codes/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, describe, part, verdict, short, range, fixturePath } from '../../lib/harness.mjs';

const FIXTURE = 'paragraph-116-codes';
const PARA = '00000002';
await header('convert every placeholder of a real 110-token paragraph (tokens separated by <w:br/>)', [FIXTURE]);

const paraXml = (docXml, id) => (docXml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${id}"[^>]*>[\\s\\S]*?</w:p>`)) || [''])[0];
// Paragraph buffer as offsets count it: w:t text, <w:br/> and <w:tab/> = one character each; pPr/rPr excluded.
function buffer(pXml) {
  let s = '';
  const inner = pXml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '');
  for (const m of inner.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:br\s*\/>|<w:tab\s*\/>/g)) s += m[1] !== undefined ? m[1] : m[0].startsWith('<w:br') ? '\n' : '\t';
  return s;
}
const tokensOf = (s) => [...s.matchAll(/\{\{[^{}]+\}\}/g)].map((m) => ({ text: m[0], s: m.index, e: m.index + m[0].length }));
const sdts = (pXml) => [...pXml.matchAll(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/g)].map((m) => buffer(`<w:p>${m[1]}</w:p>`));
const brCount = (pXml) => (pXml.match(/<w:br\s*\/>/g) || []).length;

const srcPara = paraXml(await part(fixturePath(FIXTURE), 'word/document.xml'), PARA);
const TOKENS = tokensOf(buffer(srcPara));
const BR = brCount(srcPara);
console.log(`source paragraph ${PARA}: ${TOKENS.length} tokens, ${BR} <w:br/>, first ${TOKENS[0].text} at [${TOKENS[0].s},${TOKENS[0].e}), last ${TOKENS.at(-1).text} at [${TOKENS.at(-1).s},${TOKENS.at(-1).e})`);

function inspect(bytesXml, attempted) {
  const p = paraXml(bytesXml, PARA);
  const text = buffer(p);
  const survivors = TOKENS.filter((t) => text.includes(t.text)).length;
  const wrapped = sdts(p);
  const wrong = wrapped.filter((w) => !TOKENS.some((t) => t.text === w));
  return { tokensLeft: survivors, brLeft: brCount(p), sdts: wrapped.length, wrongWraps: wrong.slice(0, 3), attempted };
}

async function pass(name, targets) {
  const { value, bytes } = await withDoc(FIXTURE, async (doc) => {
    const out = { created: 0, refused: 0, firstRefusal: null, resolved: [] };
    for (const t of targets) {
      const at = range(PARA, t.s, t.e);
      if (targets.length <= 2) {
        const probe = await attempt(() => doc.delete({ target: at, dryRun: true }));
        out.resolved.push(probe?.receipt?.resolution?.text ?? probe?.resolution?.text ?? null);
      }
      const r = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: t.text.slice(2, -2), at }));
      if (ok(r)) out.created++; else { out.refused++; out.firstRefusal ??= r.THROW ?? short(r, 160); }
    }
    return out;
  }, name);
  const res = { name, ...value, ...inspect(await part(bytes, 'word/document.xml'), targets.length) };
  res.silentLoss = res.created > 0 && (res.tokensLeft < TOKENS.length || res.brLeft < BR || res.wrongWraps.length > 0);
  console.log(`[${name}] attempted ${targets.length}: created ${res.created}, refused ${res.refused}${res.firstRefusal ? ` (${res.firstRefusal})` : ''}${res.resolved.length ? ` | dryRun resolves ${JSON.stringify(res.resolved)}` : ''}`);
  console.log(`  exported paragraph: ${res.tokensLeft}/${TOKENS.length} tokens, ${res.brLeft}/${BR} <w:br/>, ${res.sdts} SDTs${res.wrongWraps.length ? `, SDTs wrapping wrong text e.g. ${JSON.stringify(res.wrongWraps)}` : ''}${res.silentLoss ? '  <-- content lost with success:true' : ''}`);
  return res;
}

const A = await pass('first-token', [TOKENS[0]]);
const B = await pass('last-token', [TOKENS.at(-1)]);
const C = await pass('import-all', [...TOKENS].reverse());

const lossy = [A, B, C].filter((x) => x.silentLoss);
const convertedAll = C.created === TOKENS.length && !C.silentLoss;
const say = (x) => `${x.name}: created ${x.created}/${x.attempted}, ${x.tokensLeft}/${TOKENS.length} tokens and ${x.brLeft}/${BR} breaks survive${x.silentLoss ? ' (SILENT LOSS)' : ''}`;
verdict(lossy.length > 0 || !convertedAll,
  `all ${TOKENS.length} placeholders become content controls, each wrapping exactly its token, with every token and <w:br/> preserved`,
  [A, B, C].map(say).join(' | '),
  { tokens: TOKENS.length, breaks: BR, passes: [A, B, C].map(({ name, created, refused, firstRefusal, tokensLeft, brLeft, sdts, wrongWraps, silentLoss }) => ({ name, created, refused, firstRefusal, tokensLeft, brLeft, sdts, wrongWraps, silentLoss })) });
