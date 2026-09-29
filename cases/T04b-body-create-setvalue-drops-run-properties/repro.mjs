// T04b - In the BODY: wrap a styled token in a content control, then fill it with
// contentControls.text.setValue. Does the value keep the token's run formatting?
// Compared with: the mint alone (no setValue), and replace() alone on the same range.
// Each variant runs on a fresh open; the verdict comes from the exported word/document.xml.
//
// Run: node cases/T04b-body-create-setvalue-drops-run-properties/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { readFile } from 'node:fs/promises';
import { header, withDoc, attempt, describe, ok, part, findSdt, rPrOf, verdict, fixturePath, range, sdtInline } from '../../lib/harness.mjs';

const FIX = 'styled-token-body';
const P = '30000001';            // "Retention: {{fancy_field}} of the sum."
const TOKEN = '{{fancy_field}}';
const S = 11, E = S + TOKEN.length;
const VALUE = 'Fancy Value';
await header('body: create.contentControl + text.setValue vs replace() - is the run formatting kept?', [FIX]);

const src = await part(await readFile(fixturePath(FIX)), 'word/document.xml');
const before = rPrOf([...src.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((m) => m[0]).find((x) => x.includes(TOKEN)));
// rPr of the run whose text contains `needle` (outside or inside an SDT)
const runRPr = (xml, needle) => { const run = [...xml.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((m) => m[0]).find((x) => x.includes(needle)); return run ? rPrOf(run) : '(text not found)'; };

const create = (doc) => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: 'fancy_field', alias: 'Fancy Field', at: range(P, S, E) });
const VARIANTS = [
  ['create.contentControl, then contentControls.text.setValue', async (doc) => {
    const c = await attempt(() => create(doc));
    if (!ok(c)) return { steps: [['create', c]] };
    const target = (await doc.contentControls.selectByTag({ tag: 'fancy_field' })).items[0]?.target ?? sdtInline(c.contentControl?.nodeId);
    const s = await attempt(() => doc.contentControls.text.setValue({ target, value: VALUE }));
    return { steps: [['create', c], ['text.setValue', s]] };
  }, VALUE],
  ['create.contentControl only (no value written)', async (doc) => ({ steps: [['create', await attempt(() => create(doc))]] }), TOKEN],
  ['replace() only (no content control)', async (doc) => ({ steps: [['replace', await attempt(() => doc.replace({ target: range(P, S, E), text: VALUE }))]] }), VALUE],
];

console.log(`rPr of the token run in the input: ${before}`);
const rows = [];
let i = 0;
for (const [label, run, expectText] of VARIANTS) {
  const { value, bytes } = await withDoc(FIX, run, `variant-${++i}`);
  const xml = await part(bytes, 'word/document.xml');
  const sdt = findSdt(xml, 'fancy_field');
  const after = runRPr(sdt ? sdt.content : xml, expectText);
  const allOk = value.steps.every(([, r]) => ok(r));
  rows.push({ label, allOk, kept: after === before, after, sdt: !!sdt });
  console.log(label);
  for (const [name, r] of value.steps) console.log(`   ${name}: ${describe(r)}`);
  console.log(`   SDT in export: ${sdt ? 'yes' : 'no'}   rPr of "${expectText}" in export: ${after}   -> ${after === before ? 'KEPT' : 'LOST'}`);
}

const [withValue, mintOnly, replaceOnly] = rows;
verdict(withValue.allOk && !withValue.kept,
  'the value written by text.setValue keeps the wrapped token\'s <w:rPr> (Georgia, bold, italic, C00000, 12pt, pt-BR), as replace() does on the same range',
  `create + text.setValue: rPr ${withValue.kept ? 'kept' : 'LOST'}; create only: rPr ${mintOnly.kept ? 'kept' : 'LOST'}; replace() only: rPr ${replaceOnly.kept ? 'kept' : 'LOST'}`,
  { variants: rows.map(({ label, allOk, kept, after }) => ({ label, success: allOk, rPrKept: kept, rPrAfter: after })), rPrBefore: before });
