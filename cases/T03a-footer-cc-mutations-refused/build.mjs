// Builds this case's fixtures. Run: node cases/T03a-footer-cc-mutations-refused/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, tab, richSdt, plainSdt, pageField, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
writeFixtures(join(HERE, 'fixtures'), [
  // one footer paragraph with two inline rich-text SDTs; PAGE field in the sibling paragraph
  ['footer-two-rich-sdts', await buildDocx({ body, footers: [ftr(
    p('50000001', r('Seller initials: ', RPR) + richSdt('seller_initials', '41001', '[seller_initials]', RPR) + r('   ', RPR) + tab(RPR) + r('Buyer initials: ', RPR) + richSdt('buyer_initials', '41002', '[buyer_initials]', RPR))
    + p('50000002', r('Sample Co.', RPR) + tab(RPR) + pageField(RPR)))] })],
  // footer with one plain-text SDT (for contentControls.text.*)
  ['footer-plain-text-sdt', await buildDocx({ body, footers: [ftr(p('50000001', r('Initials: ', RPR) + plainSdt('initials', '41003', 'AB', RPR)))] })],
]);
