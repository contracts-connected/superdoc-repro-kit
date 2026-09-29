// Builds this case's fixtures. Run: node cases/T02a-replace-then-create-loses-siblings/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const tail = p('10000009', r('Closing paragraph.'));
writeFixtures(join(HERE, 'fixtures'), [
  // one run for the whole sentence
  ['one-run-sentence', await buildDocx({ body: p('00000001', r('Retention: {{retainage_note}} of the sum.')) + tail })],
  // three runs: bold label, the placeholder with a space on each side, the rest of the sentence
  ['three-run-sentence', await buildDocx({ body: p('00000001', r('Retention:', '<w:rPr><w:b/></w:rPr>') + r(' {{retainage_note}} ') + r('of the sum.')) + tail })],
]);
