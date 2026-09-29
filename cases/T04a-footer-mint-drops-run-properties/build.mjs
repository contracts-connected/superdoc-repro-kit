// Builds this case's fixture. Run: node cases/T04a-footer-mint-drops-run-properties/build.mjs
// Every run carries RPR = Arial, bold, colour 0B5394, size 10pt (sz 20).
//   footer 50000001  "Initials: [buyer_initials]"         token in the SAME run as its label
//   footer 50000002  "Seller: " + "[seller_initials]"     token in its OWN run
//   body   10000001  "Body initials: [body_initials]"      same-run token in the body (comparison)
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
writeFixtures(join(HERE, 'fixtures'), [
  ['styled-tokens-footer-and-body', await buildDocx({
    body: p('10000001', r('Body initials: [body_initials]', RPR)) + p('10000002', r('Lorem ipsum.')),
    footers: [ftr(p('50000001', r('Initials: [buyer_initials]', RPR)) + p('50000002', r('Seller: ', RPR) + r('[seller_initials]', RPR)))],
  })],
]);
