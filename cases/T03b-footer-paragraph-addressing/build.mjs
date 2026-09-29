// Builds this case's fixture. Run: node cases/T03b-footer-paragraph-addressing/build.mjs
// One section, titlePg.
//   footer1 (first page)  plain paragraph 50000001 "Initials: [buyer_initials]"
//   footer2 (default)     a 2x2 table; cell paragraphs 61000001..61000004, cell 1 holds "No: {{contract_number}}"
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const tc = (id, text) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${p(id, r(text))}</w:tc>`;
const table = '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
  + `<w:tr>${tc('61000001', 'No: {{contract_number}}')}${tc('61000002', 'Supplier')}</w:tr>`
  + `<w:tr>${tc('61000003', 'Initials: [seller_initials]')}${tc('61000004', 'Page')}</w:tr></w:tbl>`;
writeFixtures(join(HERE, 'fixtures'), [
  ['footer-plain-and-table', await buildDocx({
    body: p('10000001', r('Lorem ipsum dolor sit amet.')),
    footers: [ftr(p('50000001', r('Initials: [buyer_initials]'))), ftr(table + p('61000009', r('')))],
    footerRefs: [['first', 1], ['default', 2]], titlePg: true,
  })],
]);
