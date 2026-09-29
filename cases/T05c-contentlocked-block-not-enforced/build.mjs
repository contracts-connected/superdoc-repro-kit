// Builds this case's fixtures. Run: node cases/T05c-contentlocked-block-not-enforced/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, richSdt, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
writeFixtures(join(HERE, 'fixtures'), [
  // titlePg; the FIRST-PAGE footer (footer1.xml) is wrapped in a block SDT with <w:lock w:val="contentLocked"/>
  // and tag goog_rdk_0 - the wrapper Google Docs writes on export. It holds a raw token and an inline SDT.
  // The default footer (footer2.xml) is plain.
  ['first-page-footer-content-locked', await buildDocx({ body,
    footers: [
      ftr(`<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_0"/><w:id w:val="42000"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>`
        + p('50000001', r('Seller initials: [seller_initials]', RPR))
        + p('50000002', r('Buyer initials: ', RPR) + richSdt('buyer_initials', '41002', '[buyer_initials]', RPR))
        + '</w:sdtContent></w:sdt>'),
      ftr(p('50000011', r('Sample Co.', RPR))),
    ],
    footerRefs: [['first', 1], ['default', 2]], titlePg: true })],
]);
