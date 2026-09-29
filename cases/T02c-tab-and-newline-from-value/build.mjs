// Builds this case's fixture. Run: node cases/T02c-tab-and-newline-from-value/build.mjs
// One body paragraph "Address: X." - the X (offset 9..10) is the merge value we want to write.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['body-one-paragraph', await buildDocx({ body: p('20000001', r('Address: X.')) + p('20000002', r('Second paragraph.')) })],
]);
