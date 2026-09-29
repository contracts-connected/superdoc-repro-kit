// Builds this case's fixtures. Run: node cases/T09a-tracked-change-identity-and-fusion/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['review-paragraph', await buildDocx({ body: p('00000001', r('Please review the original phrase before signing.')) + p('00000002', r('Lorem ipsum dolor sit amet.')) })],
]);
