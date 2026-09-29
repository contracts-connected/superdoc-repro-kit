// T02c - Can a merge value that contains a tab ("\t") or a line break ("\n") be written as a real
// <w:tab/> / <w:br/>? Multi-line addresses and "Label<TAB>Value" pairs are ordinary merge values for us.
// Every variant runs on a fresh open; the verdict comes from the exported word/document.xml.
//
// Run: node cases/T02c-tab-and-newline-from-value/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, part, verdict, short, range, sel, pt, visible } from '../../lib/harness.mjs';

const FIX = 'body-one-paragraph';
const P = '20000001';            // "Address: X."  -> X is offset 9..10
await header('tab / line break from a value string', [FIX]);

const at = () => range(P, 9, 10);              // the "X"
const point = () => sel(pt(P, 10), pt(P, 10)); // collapsed, right after "X"
// [label, kind, char, call]   kind: 'text' = the plain-value path we need; 'html' = alternatives we tried
const VARIANTS = [
  ['replace({ text: "A\\tB" })', 'text', 'tab', (doc) => doc.replace({ target: at(), text: 'A\tB' })],
  ['replace({ text: "A\\nB" })', 'text', 'br', (doc) => doc.replace({ target: at(), text: 'A\nB' })],
  ['insert({ type: "text", value: "A\\tB" })', 'text', 'tab', (doc) => doc.insert({ target: point(), type: 'text', value: 'A\tB' })],
  ['insert({ type: "text", value: "A\\nB" })', 'text', 'br', (doc) => doc.insert({ target: point(), type: 'text', value: 'A\nB' })],
  ['insert({ type: "html", value: "A&#9;B" })', 'html', 'tab', (doc) => doc.insert({ target: point(), type: 'html', value: 'A&#9;B' })],
  ['insert({ type: "html", value: "A<span style=white-space:pre>\\t</span>B" })', 'html', 'tab', (doc) => doc.insert({ target: point(), type: 'html', value: 'A<span style="white-space:pre">\t</span>B' })],
  ['insert({ type: "html", value: "A<br>B" })', 'html', 'br', (doc) => doc.insert({ target: point(), type: 'html', value: 'A<br>B' })],
];

const paraXml = (xml, id) => (xml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${id}"[^>]*>[\\s\\S]*?</w:p>`)) || [''])[0];
const rows = [];
let i = 0;
for (const [label, kind, char, call] of VARIANTS) {
  const { value: result, bytes } = await withDoc(FIX, (doc) => attempt(() => call(doc)), `variant-${++i}`);
  const xml = await part(bytes, 'word/document.xml');
  const px = paraXml(xml, P);
  const tab = /<w:tab\/>/.test(px);
  const br = /<w:br\/>/.test(px);
  const shown = visible(px) + (br ? '   (contains <w:br/>)' : '');
  rows.push({ label, kind, char, ok: ok(result), tab, br, shown, result });
  console.log(`${label.padEnd(72)} ${ok(result) ? 'success:true' : short(result?.THROW ?? result?.failure ?? result, 70)}`);
  console.log(`${''.padEnd(72)} exported paragraph: ${JSON.stringify(shown)}  <w:tab/> ${tab ? 'YES' : 'no'} · <w:br/> ${br ? 'YES' : 'no'}`);
}

const textTab = rows.filter((x) => x.kind === 'text' && x.char === 'tab');
const textBr = rows.filter((x) => x.kind === 'text' && x.char === 'br');
const tabOk = textTab.filter((x) => x.ok && x.tab).length;
const brOk = textBr.filter((x) => x.ok && x.br).length;
const brRefused = [...new Set(textBr.filter((x) => !x.ok).map((x) => x.result?.THROW ?? short(x.result?.failure, 80)))];
const htmlBr = rows.find((x) => x.kind === 'html' && x.char === 'br');
verdict(tabOk < textTab.length || brOk < textBr.length,
  'a tab in a text value becomes a real <w:tab/> and a line break becomes <w:br/>, through both replace({ text }) and insert({ type: "text" }) - or the docs name the supported way to write them from a value',
  `tab -> <w:tab/> in ${tabOk}/${textTab.length} text calls; line break -> <w:br/> in ${brOk}/${textBr.length} text calls${brRefused.length ? ` (refused with: ${brRefused.join('; ')})` : ''}; insert html "<br>" ${htmlBr?.ok && htmlBr.br ? 'does' : 'does not'} produce <w:br/>`,
  { variants: rows.map(({ label, ok: o, tab, br, shown }) => ({ call: label, success: o, tab, br, exported: shown })) });
