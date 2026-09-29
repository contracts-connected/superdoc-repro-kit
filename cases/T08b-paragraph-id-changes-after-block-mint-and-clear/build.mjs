// Builds this case's fixtures. Run: node cases/T08b-paragraph-id-changes-after-block-mint-and-clear/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['three-paragraphs', await buildDocx({ body: p('00000001', r('Contract body text before the signature block.')) + p('00000002', r('buyer signature placeholder')) + p('00000003', r('Contract body text after the signature block.')) })],
]);
