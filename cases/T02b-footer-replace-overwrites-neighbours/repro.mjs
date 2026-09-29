// T02b - replace() in a footer story near tracked revisions / AlternateContent.
// The same token has different offsets depending on which API you ask:
//   * getText / blocks.list textPreview ("final" view)       -> tracked deletions, move-froms and AlternateContent excluded
//   * projectHtml({ includeSourceMap: true }) segment ranges -> labelled coordinateSpace "tracked"
//   * the part XML (what a caller reading word/footer1.xml counts)
// replace() accepts offsets from any of them and returns success:true, even when the range does not
// cover the token and neighbouring text is overwritten. The verdict comes from the exported footer XML.
//
// Run: node cases/T02b-footer-replace-overwrites-neighbours/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, part, verdict, short, hfStory, range } from '../../lib/harness.mjs';

const FIX = 'footer-revisions-and-alternate-content';
const VALUE = 'VALUE';
// paragraph (w14:paraId) -> token it contains, and what precedes the token
const PARAS = [
  ['1F000001', '{{project_name}}', 'control: plain runs'],
  ['1F000008', '{{owner_name}}', '<w:del> "deleted " before the token'],
  ['1F000012', '{{project_name}}', '<w:moveFrom> "moved " before the token'],
  ['1F000013', '{{project_name}}', 'mc:AlternateContent (Choice "X" / Fallback "YYY") before the token'],
];
await header('replace() offsets in a footer near tracked revisions and AlternateContent', [FIX]);

const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
// Paragraph text as a caller reading the part XML counts it: every <w:t> (not <w:delText>), pPr/rPr excluded.
function paraXml(xml, id) { const m = xml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${id}"[^>]*>([\\s\\S]*?)</w:p>`)); return m ? m[1] : null; }
const wt = (x) => [...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => decode(m[1])).join('');
// Token range in block coordinates, read back from projectHtml's sourceMap.
function sourceMapRange(ph, blockId, tok) {
  const ents = ph.sourceMap.entries.filter((e) => e.kind === 'text' && e.blockId === blockId);
  if (!ents.length) return null;
  const from = ents[0].output.start;
  const at = ph.content.slice(from, ents.at(-1).output.end).indexOf(tok);
  if (at < 0) return null;
  const toBlock = (o, isEnd) => { for (const e of ents) if (o >= e.output.start && (isEnd ? o <= e.output.end : o < e.output.end)) return e.source.segments[0].range.start + (o - e.output.start); return null; };
  return { s: toBlock(from + at, false), e: toBlock(from + at + tok.length, true), coordinateSpace: ents[0].source.coordinateSpace };
}

// 1. Read every coordinate space once, and export the untouched document as the baseline.
const { value: spaces, bytes: baseBytes } = await withDoc(FIX, async (doc) => {
  const parts = (await doc.headerFooters.parts.list({})).items;
  const story = hfStory(parts[0].refId);
  const blocks = (await doc.blocks.list({ in: story })).blocks;
  const ph = await doc.projectHtml({ in: story, includeSourceMap: true });
  const out = {};
  for (const [id, tok] of PARAS) {
    const preview = blocks.find((b) => b.nodeId === id)?.textPreview ?? '';
    const f = preview.indexOf(tok);
    out[id] = { final: f >= 0 ? { s: f, e: f + tok.length } : null, sourceMap: sourceMapRange(ph, id, tok), preview };
  }
  return { refId: parts[0].refId, getText: await doc.getText({ in: story }), out };
}, 'baseline');
const baseXml = await part(baseBytes, 'word/footer1.xml');
for (const [id, tok] of PARAS) {
  const x = wt(paraXml(baseXml, id)); const i = x.indexOf(tok);
  spaces.out[id].xml = { s: i, e: i + tok.length, text: x };
}
console.log(`getText (footer): ${JSON.stringify(spaces.getText)}`);
for (const [id, tok, what] of PARAS) {
  const o = spaces.out[id];
  console.log(`${id}  ${what}`);
  console.log(`   textPreview ${JSON.stringify(o.preview)}  -> token at ${o.final?.s}..${o.final?.e}`);
  console.log(`   sourceMap   (coordinateSpace "${o.sourceMap?.coordinateSpace}") -> token at ${o.sourceMap?.s}..${o.sourceMap?.e}`);
  console.log(`   part XML    ${JSON.stringify(o.xml.text)}  -> token at ${o.xml.s}..${o.xml.e}`);
}
console.log('-'.repeat(100));

// 2. For each paragraph and each reading: fresh open, replace(token -> VALUE), export, compare with the baseline.
const rows = [];
for (const [id, tok] of PARAS) {
  for (const space of ['final', 'sourceMap', 'xml']) {
    const rg = spaces.out[id][space];
    if (!rg || rg.s == null) { rows.push({ id, space, skipped: true }); continue; }
    const { value: result, bytes } = await withDoc(FIX, async (doc) => {
      const story = hfStory(spaces.refId);
      return attempt(() => doc.replace({ target: range(id, rg.s, rg.e, story), text: VALUE }));
    }, `${id}-${space}`);
    const got = wt(paraXml(await part(bytes, 'word/footer1.xml'), id));
    const want = spaces.out[id].xml.text.replace(tok, VALUE);
    const correct = got === want;
    rows.push({ id, space, ok: ok(result), correct, got, want, result });
    console.log(`${id} ${space.padEnd(9)} replace ${String(rg.s).padStart(2)}..${String(rg.e).padEnd(2)} -> ${ok(result) ? 'success:true ' : short(result?.THROW ?? result?.failure ?? result, 80)}  exported ${correct ? 'CORRECT' : 'WRONG'} ${correct ? '' : `got ${JSON.stringify(got)}`}`);
  }
}

// 3. Verdict: a range the host itself produced (textPreview or sourceMap) is accepted with success:true
//    but the exported footer is not "baseline with only the token replaced".
const silent = rows.filter((x) => x.ok && x.correct === false);
const hostSilent = silent.filter((x) => x.space !== 'xml');
const spacesDisagree = PARAS.some(([id]) => { const o = spaces.out[id]; return o.final?.s !== o.sourceMap?.s; });
verdict(hostSilent.length > 0 || (silent.length > 0 && spacesDisagree),
  'one documented coordinate space shared by getText/blocks.list, the projectHtml sourceMap and replace(); a replace() whose range does not cover the intended text fails (or can be guarded by an expected-text precondition) instead of overwriting neighbours',
  `${silent.length} replace() calls returned success:true and overwrote neighbouring text (${silent.map((x) => `${x.id} via ${x.space}`).join(', ')}); textPreview and sourceMap offsets ${spacesDisagree ? 'disagree' : 'agree'} for the same token`,
  { silentOverwrites: silent.map((x) => ({ paragraph: x.id, offsetsFrom: x.space, exported: x.got, expected: x.want })) });
