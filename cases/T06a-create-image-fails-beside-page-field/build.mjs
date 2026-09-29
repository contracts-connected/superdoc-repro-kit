// Builds this case's fixtures. Run: node cases/T06a-create-image-fails-beside-page-field/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, ftr, p, r, tab, richSdt, pageField, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Sed do eiusmod tempor.'));
writeFixtures(join(HERE, 'fixtures'), [
  // SDT and a PAGE complex field (w:fldChar) in the SAME footer paragraph
  ['footer-field-same-paragraph', await buildDocx({ body, footers: [ftr(
    p('50000001', r('Initials: ', RPR) + richSdt('buyer_initials', '41002', '[buyer_initials]', RPR, 'sdtLocked') + tab(RPR) + pageField(RPR)))] })],
  // Control: identical content, PAGE field moved to a sibling paragraph
  ['footer-field-sibling-paragraph', await buildDocx({ body, footers: [ftr(
    p('50000001', r('Initials: ', RPR) + richSdt('buyer_initials', '41002', '[buyer_initials]', RPR, 'sdtLocked'))
    + p('50000002', pageField(RPR)))] })],
]);
