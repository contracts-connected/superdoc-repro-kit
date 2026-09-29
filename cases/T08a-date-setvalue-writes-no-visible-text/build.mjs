// Builds this case's fixtures. Run: node cases/T08a-date-setvalue-writes-no-visible-text/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['signed-on-token', await buildDocx({ body: p('00000001', r('Signed on: {{sig_date}}')) + p('00000002', r('Lorem ipsum dolor sit amet.')) })],
]);
