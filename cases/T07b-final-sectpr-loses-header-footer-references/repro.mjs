// T07b - open + save of an UNCHANGED multi-section document: does every <w:sectPr> keep its header/footer references?
// We saw the FINAL body-level sectPr lose its default headerReference/footerReference on a real two-section template
// (titlePg in section 1). References are compared by type and by the part they point at (r:id values may be renumbered).
//
// Run: node cases/T07b-final-sectpr-loses-header-footer-references/repro.mjs [2.13.0|2.15.0|2.16.0-next.17]
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { header, withDoc, part, partNames, fixturePath, verdict, ROOT } from '../../lib/harness.mjs';

const FIXTURES = ['two-sections-final-has-content', 'two-sections-final-empty'];
if (existsSync(join(ROOT, 'fixtures', 'sanitized', 'po-locked-footer.docx'))) FIXTURES.push('po-locked-footer');
await header('unchanged open + save: header/footer references of every sectPr', FIXTURES);

async function sectRefs(bytes) {
  const doc = await part(bytes, 'word/document.xml');
  const rels = await part(bytes, 'word/_rels/document.xml.rels');
  const target = Object.fromEntries([...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => [m[0].match(/Id="([^"]+)"/)[1], m[0].match(/Target="([^"]+)"/)[1]]));
  const sects = [...doc.matchAll(/<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/g)].map((m) => m[1]);
  return sects.map((s) => ({
    refs: [...s.matchAll(/<w:(header|footer)Reference\b[^>]*\/>/g)].map((m) => `${m[1]}:${m[0].match(/w:type="(\w+)"/)?.[1]}->${target[m[0].match(/r:id="([^"]+)"/)?.[1]] ?? '(missing rel)'}`).sort(),
    titlePg: /<w:titlePg\b(?![^>]*w:val="(0|false)")/.test(s),
  }));
}

const lost = [];
for (const f of FIXTURES) {
  const before = await sectRefs(await readFile(fixturePath(f)));
  const { bytes } = await withDoc(f, async () => null, `unchanged-${f}`);
  const after = await sectRefs(bytes);
  const parts = (await partNames(bytes)).filter((n) => /^word\/(header|footer)\d+\.xml$/.test(n)).sort();
  console.log(`\n${f}: ${before.length} sectPr before, ${after.length} after; header/footer parts in the saved package: ${parts.join(', ')}`);
  before.forEach((b, i) => {
    const a = after[i] ?? { refs: [], titlePg: false };
    const missing = b.refs.filter((x) => !a.refs.some((y) => y.split('->')[0] === x.split('->')[0]));
    const label = i === before.length - 1 ? 'final (body-level) sectPr' : `sectPr #${i + 1} (paragraph-level)`;
    console.log(`  ${label}`);
    console.log(`    before: ${b.refs.join(', ') || '(no references)'}${b.titlePg ? ' + titlePg' : ''}`);
    console.log(`    after:  ${a.refs.join(', ') || '(no references)'}${a.titlePg ? ' + titlePg' : ''}`);
    if (missing.length || b.titlePg !== a.titlePg) {
      lost.push(`${f} ${label}: ${missing.join(', ')}${b.titlePg !== a.titlePg ? ' titlePg' : ''}`);
      console.log(`    LOST:   ${missing.join(', ')}${b.titlePg !== a.titlePg ? ' titlePg' : ''}`);
    }
  });
}

verdict(lost.length > 0,
  'an unchanged open + save keeps every sectPr with the same header/footer references (same type, same target part) and titlePg',
  lost.length ? `references lost: ${lost.join(' | ')}` : `every sectPr kept its header/footer references on ${FIXTURES.length} documents`,
  { lost });
