// Builds this case's fixture. Run: node cases/T02b-footer-replace-overwrites-neighbours/build.mjs
// One footer, four paragraphs, each holding one {{token}}:
//   1F000001  control: plain runs only
//   1F000008  a tracked deletion (<w:del>) before the token
//   1F000012  a tracked move-from (<w:moveFrom>) before the token
//   1F000013  an mc:AlternateContent run (Choice "X", Fallback "YYY") before the token
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const RI = '<w:rPr><w:i/><w:sz w:val="20"/></w:rPr>';
const REV = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const footer = ftr(
  p('1F000001', r('A{{project_name}}A plain control paragraph', RI))
  + p('1F000008', `<w:del w:id="12" ${REV}><w:r><w:delText xml:space="preserve">deleted </w:delText></w:r></w:del>` + r('H{{owner_name}}H and a tail', RI))
  + p('1F000012', `<w:moveFrom w:id="13" ${REV}><w:r><w:t xml:space="preserve">moved </w:t></w:r></w:moveFrom>` + r('U{{project_name}}U tail long enough text', RI))
  + p('1F000013', r('V', RI) + '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>X</w:t></mc:Choice><mc:Fallback><w:t>YYY</w:t></mc:Fallback></mc:AlternateContent></w:r>' + r('W{{project_name}}W and a long tail of text here', RI)),
);
writeFixtures(join(HERE, 'fixtures'), [
  ['footer-revisions-and-alternate-content', await buildDocx({ body: p('10000001', r('Lorem ipsum dolor sit amet.')), footers: [footer] })],
]);
