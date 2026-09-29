// Builds this case's fixtures. Run: node cases/T10c-metadata-attach-inside-content-control/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r, richSdt } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['inline-sdt-in-body', await buildDocx({ body: p('00000001', r('Plain body text outside any control.')) + p('00000002', r('Buyer initials: ') + richSdt('buyer_initials', '41002', 'AB')) })],
]);
