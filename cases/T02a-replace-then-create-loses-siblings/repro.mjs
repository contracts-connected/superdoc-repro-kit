// T02a - a range reported by replace() is not safe to hand to create.contentControl.
//   "replace-first": replace({{retainage_note}} -> "Retainage Note"), then create.contentControl on the range the
//                    replace receipt reports (receipt.effects.insertedText[0].selectionTarget, else
//                    receipt.resolution.selectionTarget). Sibling text in the paragraph is lost.
//   "create-first" (control): create.contentControl on the untouched token range, then text.setValue. Clean.
// Paragraph text "Retention: {{retainage_note}} of the sum.", token at [11,29) in paragraph 00000001.
//
// Run: node cases/T02a-replace-then-create-loses-siblings/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, describe, part, verdict, short, range } from '../../lib/harness.mjs';

const FIXTURES = ['one-run-sentence', 'three-run-sentence'];
const PARA = '00000001';
const TAG = 'retainage_note';
const VALUE = 'Retainage Note';
const EXPECTED = 'Retention: Retainage Note of the sum.';
await header('replace() then create.contentControl on the range replace() reported', FIXTURES);

const paraXml = (docXml) => (docXml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${PARA}"[^>]*>[\\s\\S]*?</w:p>`)) || docXml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/) || [''])[0];
const textOf = (xml) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
const runs = (xml) => (xml.match(/<w:r>|<w:r\s[^>]*>/g) || []).length;
const firstRunBold = (xml) => /<w:b\s*\/>/.test((xml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/) || [''])[0]);
const receiptOf = (res) => res?.receipt ?? res;

async function replaceFirst(fixture) {
  const { value, bytes } = await withDoc(fixture, async (doc) => {
    const replaced = await attempt(() => doc.replace({ target: range(PARA, 11, 29), text: VALUE }));
    const rc = receiptOf(replaced);
    let fresh = rc?.effects?.insertedText?.[0]?.selectionTarget ?? rc?.resolution?.selectionTarget ?? null;
    let freshFrom = fresh ? 'replace receipt' : null;
    if (!fresh && ok(replaced)) {
      // this host's receipt carries no range: re-locate the text replace() just wrote
      const m = await attempt(() => doc.query.match({ in: { kind: 'story', storyType: 'body' }, select: { type: 'text', pattern: VALUE, mode: 'contains' } }));
      fresh = m?.items?.[0]?.target ?? null;
      freshFrom = fresh ? 'query.match re-locate' : null;
    }
    const created = fresh ? await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: TAG, alias: VALUE, at: fresh })) : { THROW: 'no range to wrap' };
    return { replaced, fresh, freshFrom, created };
  }, `${fixture}-replace-first`);
  const p = paraXml(await part(bytes, 'word/document.xml'));
  return { ...value, text: textOf(p), runs: runs(p), bold: firstRunBold(p), xml: p };
}

async function createFirst(fixture) {
  const { value, bytes } = await withDoc(fixture, async (doc) => {
    const created = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: TAG, alias: VALUE, at: range(PARA, 11, 29) }));
    const sel = await attempt(() => doc.contentControls.selectByTag({ tag: TAG }));
    const setValue = sel?.items?.[0] ? await attempt(() => doc.contentControls.text.setValue({ target: sel.items[0].target, value: VALUE })) : { THROW: 'control not found' };
    return { created, setValue };
  }, `${fixture}-create-first`);
  const p = paraXml(await part(bytes, 'word/document.xml'));
  return { ...value, text: textOf(p), runs: runs(p), bold: firstRunBold(p), xml: p };
}

const rows = [];
for (const fixture of FIXTURES) {
  const bad = await replaceFirst(fixture);
  const good = await createFirst(fixture);
  const wantBold = fixture === 'three-run-sentence';
  console.log(`[${fixture}]`);
  console.log(`  replace-first: replace ${bad.replaced?.THROW ? describe(bad.replaced) : `receipt.success:${receiptOf(bad.replaced)?.success}`} | range from ${bad.freshFrom ?? 'nowhere'}: ${short(bad.fresh, 200)} | create ${describe(bad.created)}`);
  console.log(`     exported text: ${JSON.stringify(bad.text)}  runs=${bad.runs}${wantBold ? ` boldLabel=${bad.bold}` : ''}`);
  console.log(`     exported XML:  ${short(bad.xml, 700)}`);
  console.log(`  create-first:  create ${describe(good.created)} | setValue ${describe(good.setValue)}`);
  console.log(`     exported text: ${JSON.stringify(good.text)}  runs=${good.runs}${wantBold ? ` boldLabel=${good.bold}` : ''}`);
  const loss = ok(bad.created) && (bad.text !== EXPECTED || (wantBold && !bad.bold));
  const controlClean = ok(good.created) && ok(good.setValue) && good.text === EXPECTED && (!wantBold || good.bold);
  rows.push({ fixture, replaceFirst: { created: ok(bad.created), text: bad.text, freshFrom: bad.freshFrom, loss }, createFirst: { clean: controlClean, text: good.text } });
}

verdict(rows.some((x) => x.replaceFirst.loss),
  `after replace() + create.contentControl on the reported range the paragraph reads "${EXPECTED}" (same as the create-first sequence)`,
  rows.map((x) => `${x.fixture}: replace-first ${x.replaceFirst.created ? `-> ${JSON.stringify(x.replaceFirst.text)}${x.replaceFirst.loss ? ' (LOST TEXT)' : ''}` : 'create failed'}; create-first -> ${JSON.stringify(x.createFirst.text)}${x.createFirst.clean ? ' (clean)' : ' (NOT clean)'}`).join(' | '),
  { rows });
