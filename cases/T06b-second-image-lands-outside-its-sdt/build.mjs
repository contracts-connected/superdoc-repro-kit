// Builds this case's fixtures. Run: node cases/T06b-second-image-lands-outside-its-sdt/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, tab, richSdt, pageField, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
writeFixtures(join(HERE, 'fixtures'), [
  // One footer paragraph with two inline rich-text SDTs (sdtLocked, as in our real templates);
  // the PAGE field sits in the sibling paragraph so it cannot interfere (see T06a).
  ['footer-two-initials-sdts', await buildDocx({ body, footers: [ftr(
    p('50000001', r('Seller initials: ', RPR) + richSdt('seller_initials', '41001', '[seller_initials]', RPR, 'sdtLocked') + r('   ', RPR) + tab(RPR) + r('Buyer initials: ', RPR) + richSdt('buyer_initials', '41002', '[buyer_initials]', RPR, 'sdtLocked'))
    + p('50000002', r('Sample Co.', RPR) + tab(RPR) + pageField(RPR)))] })],
]);
