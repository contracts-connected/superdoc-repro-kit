// T01b - create.contentControl on a placeholder that shares ONE <w:r> with a <w:tab/>.
//   Shape 1: <w:r><w:t>{{f1}}</w:t><w:tab/><w:t>{{f2}}</w:t></w:r>, wrap {{f1}} - the token BEFORE the tab (offsets 0..6)
//   Shape 2: <w:r><w:t>Line one prose.</w:t><w:tab/><w:t>{{t}} more prose.</w:t></w:r>, wrap {{t}} - the token AFTER the tab
// Same failure family as T01a, with a tab instead of a hard line break.
//
// Run: node cases/T01b-create-before-tab-deletes-tab/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, describe, part, verdict, short, range } from '../../lib/harness.mjs';

await header('create.contentControl next to a <w:tab/> in the same run deletes the tab and the next token', ['one-run-two-tokens-tab', 'one-run-prose-tab-token']);

const paraXml = (docXml, id) => (docXml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${id}"[^>]*>[\\s\\S]*?</w:p>`)) || docXml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/) || [''])[0];
const flat = (xml) => {
  let s = '';
  for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:br\/>|<w:tab\/>|<w:sdtContent>|<\/w:sdtContent>/g)) {
    s += m[1] !== undefined ? m[1] : m[0] === '<w:br/>' ? '<BR>' : m[0] === '<w:tab/>' ? '<TAB>' : m[0] === '<w:sdtContent>' ? '[[' : ']]';
  }
  return s;
};
const sdtText = (xml) => { const m = xml.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/); return m ? flat(m[1]) : null; };

async function shape(fixture, paraId, s, e, tag, expect) {
  const target = range(paraId, s, e);
  const { value, bytes } = await withDoc(fixture, async (doc) => {
    const probe = await attempt(() => doc.delete({ target, dryRun: true }));
    const resolvedText = probe?.receipt?.resolution?.text ?? probe?.resolution?.text ?? null;
    const created = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag, at: target }));
    return { resolvedText, created };
  }, `${fixture}-${tag}`);
  const para = paraXml(await part(bytes, 'word/document.xml'), paraId);
  const after = flat(para);
  const wrapped = sdtText(para);
  console.log(`[${fixture}] target ${paraId} [${s},${e}) resolves to ${JSON.stringify(value.resolvedText)} (non-mutating dryRun read)`);
  console.log(`  create.contentControl: ${describe(value.created)} ${ok(value.created) ? short(value.created.contentControl ?? value.created, 160) : ''}`);
  console.log(`  paragraph before: ${expect.before}`);
  console.log(`  paragraph after:  ${after}`);
  console.log(`  exported XML:     ${short(para, 600)}`);
  const lost = expect.mustSurvive.filter((piece) => !after.includes(piece));
  const refused = !ok(value.created);
  const wrongWrap = !refused && wrapped !== expect.wrap;
  console.log(`  lost: ${lost.length ? lost.join(', ') : 'nothing'} | SDT wraps: ${JSON.stringify(wrapped)}${wrongWrap ? ' (WRONG)' : ''}${refused ? ' | refused cleanly' : ''}`);
  return { fixture, refused, lost, wrapped, wrongWrap, silentLoss: !refused && (lost.length > 0 || wrongWrap) };
}

const a = await shape('one-run-two-tokens-tab', '00000006', 0, 6, 'f1',
  { before: '{{f1}}<TAB>{{f2}}', mustSurvive: ['<TAB>', '{{f2}}'], wrap: '{{f1}}' });
const b = await shape('one-run-prose-tab-token', '0000000D', 16, 21, 't',
  { before: 'Line one prose.<TAB>{{t}} more prose.', mustSurvive: ['Line one prose.', '<TAB>', ' more prose.'], wrap: '{{t}}' });

const say = (x) => x.refused ? 'refused (no loss)' : x.silentLoss ? `success:true but lost [${x.lost.join(', ')}]${x.wrongWrap ? ` and SDT wraps ${JSON.stringify(x.wrapped)}` : ''}` : 'clean';
verdict(a.silentLoss || b.silentLoss,
  'the SDT wraps exactly the target token; the <w:tab/> and every sibling text node survive (or the call fails with a failure object instead of success:true)',
  `token before tab: ${say(a)} | token after tab: ${say(b)}`,
  { beforeTab: a, afterTab: b });
