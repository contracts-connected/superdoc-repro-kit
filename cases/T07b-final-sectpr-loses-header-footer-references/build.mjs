// Builds this case's fixtures. Run: node cases/T07b-final-sectpr-loses-header-footer-references/build.mjs
// Two sections. Section 1 (paragraph-level sectPr): titlePg + first/default header and footer.
// Section 2 (final, body-level sectPr): default header and footer.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, hdr, ftr, p, r, PG } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const headers = [hdr(p('60000001', r('Header first page'))), hdr(p('60000002', r('Header section 1 default'))), hdr(p('60000003', r('Header section 2 default')))];
const footers = [ftr(p('50000001', r('Footer first page'))), ftr(p('50000002', r('Footer section 1 default'))), ftr(p('50000003', r('Footer section 2 default')))];
const sect1 = `<w:pPr><w:sectPr><w:headerReference w:type="first" r:id="rIdH1"/><w:headerReference w:type="default" r:id="rIdH2"/><w:footerReference w:type="first" r:id="rIdF1"/><w:footerReference w:type="default" r:id="rIdF2"/>${PG}<w:titlePg/></w:sectPr></w:pPr>`;
const finalSect = `<w:sectPr><w:headerReference w:type="default" r:id="rIdH3"/><w:footerReference w:type="default" r:id="rIdF3"/>${PG}</w:sectPr>`;
const s1 = p('10000001', r('Section one, first paragraph.')) + p('10000002', r('Section one, last paragraph.'), sect1);
writeFixtures(join(HERE, 'fixtures'), [
  ['two-sections-final-has-content', await buildDocx({ body: s1 + p('10000003', r('Section two paragraph.')), headers, footers, finalSectPr: finalSect })],
  ['two-sections-final-empty', await buildDocx({ body: s1, headers, footers, finalSectPr: finalSect })],
]);
