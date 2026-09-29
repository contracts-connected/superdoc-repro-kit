// Builds this case's fixtures. Run: node cases/T01c-footer-mint-after-tab-refused/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, tab, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
const TABS = '<w:pPr><w:tabs><w:tab w:val="left" w:pos="6480"/></w:tabs></w:pPr>';
writeFixtures(join(HERE, 'fixtures'), [
  // footer: "Initials:" <tab> "[buyer_initials]" - three runs, same rPr
  ['footer-token-after-tab', await buildDocx({ body, footers: [ftr(p('50000001', r('Initials:', RPR) + tab(RPR) + r('[buyer_initials]', RPR)))] })],
  // footer laid out like our work-order template: a paragraph tab stop, two labels, a run-level tab between them
  ['footer-two-labels-tab-stop', await buildDocx({ body, footers: [ftr(p('50000001',
    r('Subcontractor Initials: [seller_initials]', RPR) + r('          ', RPR) + tab(RPR) + r('Contractor Initials: [buyer_initials]', RPR) + r(' tail', RPR), TABS))] })],
]);
