// T08a - contentControls.date.setValue stores the date (w:date/@w:fullDate) but writes no visible text into
// <w:sdtContent>, so the control keeps showing whatever it wrapped (here the raw token). text.setValue on a text
// control, the control case, does replace the visible text.
// Paragraph 00000001: "Signed on: {{sig_date}}" - the token is offsets 11..23.
//
// Run: node cases/T08a-date-setvalue-writes-no-visible-text/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, describe, ok, part, findSdt, visible, verdict, range, sdtInline, short, hasFn } from '../../lib/harness.mjs';

const FIX = 'signed-on-token';
await header('date.setValue vs text.setValue: is the value visible in the exported sdtContent?', [FIX]);

async function run(controlType) {
  const tag = `sig_date_${controlType}`;
  const { value, bytes } = await withDoc(FIX, async (doc) => {
    const created = await attempt(() => doc.create.contentControl({ kind: 'inline', controlType, tag, alias: tag, at: range('00000001', 11, 23) }));
    const nodeId = created?.contentControl?.nodeId;
    const target = sdtInline(nodeId);
    const set = controlType === 'date'
      ? await attempt(() => doc.contentControls.date.setValue({ target, value: '2026-09-21' }))
      : await attempt(() => doc.contentControls.text.setValue({ target, value: '2026-09-21' }));
    const got = hasFn(doc, 'contentControls.get') ? await attempt(() => doc.contentControls.get({ target })) : null;
    return { created, set, got };
  }, `${controlType}-control`);
  const xml = await part(bytes, 'word/document.xml');
  const sdt = findSdt(xml, tag);
  const shown = sdt ? visible(sdt.content) : null;
  const fullDate = sdt?.pr.match(/w:fullDate="([^"]+)"/)?.[1] ?? null;
  console.log(`${controlType} control:`);
  console.log(`  create.contentControl ${describe(value.created)} nodeId=${value.created?.contentControl?.nodeId}`);
  console.log(`  ${controlType}.setValue('2026-09-21') ${describe(value.set)}`);
  if (value.got) console.log(`  contentControls.get -> ${short(value.got, 260)}`);
  console.log(`  exported sdtPr date: ${fullDate ?? '(none)'} | exported visible sdtContent: ${JSON.stringify(shown)}`);
  console.log(`  exported SDT XML: ${sdt ? short(sdt.xml, 500) : '(no SDT in export)'}`);
  return { ok: ok(value.created) && ok(value.set), shown, fullDate };
}

const date = await run('date');
const text = await run('text');
const dateVisible = date.shown != null && /2026|Sep|21/.test(date.shown);
const textVisible = text.shown != null && text.shown.includes('2026-09-21');
verdict(date.ok && !dateVisible && textVisible,
  'after date.setValue the control shows the date in its sdtContent (as text.setValue does for a text control)',
  `date control: setValue ${date.ok ? 'success' : 'failed'}, w:fullDate=${date.fullDate ?? 'none'}, visible text ${JSON.stringify(date.shown)} | text control: visible text ${JSON.stringify(text.shown)}`,
  { date, text });
