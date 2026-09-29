// T10b - text inside a text box (an inline drawing whose wps:txbx holds a paragraph) is invisible to every read
// path we found: getText, blocks.list, query.match; there is no textbox namespace. Writing a value into it (a merge
// field) is therefore impossible through the Document API.
// Body: "Body text before the text box." / [drawing with text box "Textbox content {{project_name}}"] / "Body text after..."
//
// Run: node cases/T10b-textbox-content-unreachable/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, part, verdict, BODY, range, describe, ok, short } from '../../lib/harness.mjs';

const FIX = 'textbox';
const NEEDLE = 'Textbox content';
await header('reading and writing text inside a text box (wps:txbx)', [FIX]);

const { value, bytes } = await withDoc(FIX, async (doc) => {
  const v = {};
  const text = await attempt(() => doc.getText({}));
  v.getText = typeof text === 'string' ? text : text?.text ?? text;
  v.getTextHas = String(v.getText ?? '').includes(NEEDLE);
  const blocks = await attempt(() => doc.blocks.list({}));
  v.blocks = (blocks?.blocks ?? []).map((b) => `${b.nodeType}:${b.nodeId}:"${(b.textPreview ?? '').slice(0, 30)}"`);
  v.blocksHave = JSON.stringify(blocks).includes(NEEDLE);
  const m1 = await attempt(() => doc.query.match({ in: BODY, select: { type: 'text', pattern: NEEDLE, mode: 'contains' } }));
  const m2 = await attempt(() => doc.query.match({ select: { type: 'text', pattern: NEEDLE, mode: 'contains' } }));
  v.match = [m1?.THROW ?? m1?.total, m2?.THROW ?? m2?.total];
  const hit = m2?.items?.[0];
  v.hitTarget = hit ? short(hit.target ?? hit, 400) : null;
  // write through the match's own target (the only address we get for text box text)
  const tgt = hit?.target ?? hit?.selectionTarget ?? hit?.handle;
  v.replaceViaMatch = tgt ? await attempt(() => doc.replace({ target: tgt, text: 'Replaced via match target' })) : { THROW: 'no match target' };
  v.textboxApi = Object.keys(doc).filter((k) => /text.?box|shape|drawing/i.test(k));
  // try to write through the text box paragraph's own id (w14:paraId 00000009)
  v.replaceInBox = await attempt(() => doc.replace({ target: range('00000009', 16, 32), text: 'Project Alpha' }));
  console.log(`getText: ${JSON.stringify(v.getText)} (contains "${NEEDLE}": ${v.getTextHas})`);
  console.log(`blocks.list: ${v.blocks.join(' | ')} (any block mentions it: ${v.blocksHave})`);
  console.log(`query.match "${NEEDLE}": in:body -> ${v.match[0]} | no \`in\` -> ${v.match[1]}`);
  console.log(`  match item target: ${v.hitTarget}`);
  console.log(`replace using that match target -> ${describe(v.replaceViaMatch)}`);
  console.log(`document namespaces that look like text box / shape APIs: ${v.textboxApi.join(', ') || '(none)'}`);
  console.log(`replace on the text box paragraph id 00000009 -> ${describe(v.replaceInBox)}`);
  return v;
}, 'after-replace-attempt');
const xml = await part(bytes, 'word/document.xml');
const boxText = xml.match(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/)?.[1]?.match(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)?.join('') ?? '';
const written = boxText.includes('Project Alpha') || boxText.includes('Replaced via match target');
console.log(`exported text box content: ${short(boxText.replace(/<[^>]+>/g, ''), 120)}`);

const readable = value.getTextHas || value.blocksHave || value.match.some((x) => typeof x === 'number' && x > 0);
verdict(!written,
  'text box content is readable (getText / blocks.list / query.match) and writable (replace on a returned target, or a textbox API) like body text',
  `getText has it: ${value.getTextHas}; blocks.list has it: ${value.blocksHave}; query.match totals: ${value.match.join(' / ')}; replace by paragraph id: ${describe(value.replaceInBox)}; replace via the match target: ${describe(value.replaceViaMatch)}; exported box written: ${written}`,
  { readable, written });
