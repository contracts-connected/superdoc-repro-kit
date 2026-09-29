// T01c - mint an inline rich-text content control over a raw footer token that follows a <w:tab/>.
// Offsets come from the host itself: the line of getText({ in: footerStory }) that holds the token. For
// comparison the script also mints with offsets counted from the part XML (w:t characters + each run-level
// <w:tab/> as one character). Each mint runs on a fresh open; the result is checked in word/footer1.xml.
//
// Run: node cases/T01c-footer-mint-after-tab-refused/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { header, withDoc, attempt, ok, describe, part, verdict, short, hasFn, findSdt, visible, rPrOf, pt, sel, fixturePath } from '../../lib/harness.mjs';

const CASES = [
  ['footer-token-after-tab', '[buyer_initials]', 'buyer_initials'],
  ['footer-two-labels-tab-stop', '[buyer_initials]', 'buyer_initials'],
  // sanitized real template (fixtures/sanitized/): the layout where we first saw a one-character skew
  ['work-order-footer-initials', '[buyer_initials]', 'buyer_initials'],
];
await header('mint an inline SDT over a raw footer token that follows a <w:tab/>', CASES.map((c) => c[0]));

// The paragraph of footer1.xml that contains `token`: its w14:paraId and its text counted from the XML
// (w:t text + run-level <w:tab/>/<w:br/> = 1 character, pPr excluded).
function xmlParagraph(footerXml, token) {
  for (const m of footerXml.matchAll(/<w:p\b([^>]*)>([\s\S]*?)<\/w:p>/g)) {
    const inner = m[2].replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '');
    let text = '';
    for (const t of inner.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/>/g)) text += t[1] !== undefined ? t[1] : '\t';
    if (text.includes(token)) return { paraId: (m[1].match(/w14:paraId="([0-9A-F]+)"/) || [])[1], text, xmlStart: text.indexOf(token) };
  }
  throw new Error(`token ${token} not found in footer1.xml`);
}

async function mint(fixture, token, tag, para, how) {
  const { value, bytes } = await withDoc(fixture, async (doc) => {
    const parts = (await doc.headerFooters.parts.list({})).items;
    const refId = (parts.find((x) => x.partPath === 'word/footer1.xml') ?? parts[0]).refId;
    const story = { kind: 'story', storyType: 'headerFooterPart', refId };
    const text = await doc.getText({ in: story });
    // paragraph-local offset as the host sees it: the host's line that contains the token
    const line = text.split(/\r?\n/).find((l) => l.includes(token)) ?? text;
    const hostStart = line.indexOf(token);
    const start = how === 'host' ? hostStart : para.xmlStart;
    const at = sel(pt(para.paraId, start, story), pt(para.paraId, start + token.length, story));
    const input = { kind: 'inline', controlType: 'richText', tag, alias: tag, at };
    const resolve = hasFn(doc, 'capabilities.resolve') ? await attempt(() => doc.capabilities.resolve({ operationId: 'create.contentControl', input })) : null;
    const created = await attempt(() => doc.create.contentControl(input));
    return { text: line, hostStart, start, resolve, created };
  }, `${fixture}-${how}-offsets`);
  const xml = await part(bytes, 'word/footer1.xml');
  const sdt = findSdt(xml, tag);
  return { ...value, footer: visible(xml), wraps: sdt ? visible(sdt.content) : null, rPrInside: sdt ? rPrOf(sdt.content) : null };
}

const says = (v) => (v.resolve == null ? 'n/a' : v.resolve.THROW ? 'THROWS' : v.resolve.kind ?? v.resolve.available?.kind ?? short(v.resolve, 60));
const rows = [];
for (const [fixture, token, tag] of CASES) {
  const para = xmlParagraph(await part(fixturePath(fixture), 'word/footer1.xml'), token);
  const h = await mint(fixture, token, tag, para, 'host');
  const x = await mint(fixture, token, tag, para, 'xml');
  console.log(`[${fixture}] paragraph ${para.paraId} host text=${JSON.stringify(h.text)}`);
  console.log(`  host offsets [${h.start},${h.start + token.length})  resolve=${says(h)}  create: ${describe(h.created)}  SDT wraps ${JSON.stringify(h.wraps)}`);
  console.log(`  xml  offsets [${para.xmlStart},${para.xmlStart + token.length})  resolve=${says(x)}  create: ${describe(x.created)}  SDT wraps ${JSON.stringify(x.wraps)}`);
  console.log(`  exported footer (host offsets): ${short(h.footer, 400)}`);
  if (ok(h.created)) console.log(`  rPr inside the new SDT (host offsets): ${h.rPrInside}`);
  // A package-level rejection (the content-control identity scan refusing the package) is a different
  // problem (see T07a) - report it, but do not count it as the after-tab behaviour this case is about.
  const packageRejected = !ok(h.created) && /malformed package XML/i.test(h.created.THROW ?? JSON.stringify(h.created));
  if (packageRejected) console.log('  NOTE: the whole package was rejected by the content-control identity scan (see T07a); not counted here');
  const refused = !ok(h.created) && !packageRejected;
  const misWrap = ok(h.created) && h.wraps !== token;
  const xmlSkew = ok(x.created) && x.wraps !== token ? x.wraps : null;
  rows.push({ fixture, refused, misWrap, packageRejected, wraps: h.wraps, resolve: says(h), refusal: !ok(h.created) ? (h.created.THROW ?? short(h.created, 160)) : null, hostStart: h.start, xmlStart: para.xmlStart, xmlSkew });
}

const bad = rows.filter((x) => x.refused || x.misWrap);
verdict(bad.length > 0,
  'create.contentControl succeeds over a footer token that follows a <w:tab/>, and the new SDT wraps exactly the token',
  rows.map((x) => `${x.fixture}: ${x.packageRejected ? `package rejected (${x.refusal}) - see T07a` : x.refused ? `REFUSED (${x.refusal}; resolve=${x.resolve})` : x.misWrap ? `created but wraps ${JSON.stringify(x.wraps)}` : 'created, wraps the token'}${x.hostStart !== x.xmlStart ? ` [host offset ${x.hostStart} vs XML offset ${x.xmlStart}]` : ''}${x.xmlSkew ? ` [XML-offset mint wraps ${JSON.stringify(x.xmlSkew)}]` : ''}`).join(' | '),
  { rows });
