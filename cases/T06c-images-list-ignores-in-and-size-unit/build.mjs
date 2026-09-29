// Builds this case's fixtures. Run: node cases/T06c-images-list-ignores-in-and-size-unit/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, tab, richSdt, pageField, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
writeFixtures(join(HERE, 'fixtures'), [
  // Body with plain paragraphs (10000001, 10000002); footer with an inline rich-text SDT and no image.
  ['footer-initials-sdt', await buildDocx({ body, footers: [ftr(
    p('50000001', r('Seller initials: ', RPR) + richSdt('seller_initials', '41001', '[seller_initials]', RPR))
    + p('50000002', r('Sample Co.', RPR) + tab(RPR) + pageField(RPR)))] })],
]);
