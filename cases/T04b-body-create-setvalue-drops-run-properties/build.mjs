// Builds this case's fixture. Run: node cases/T04b-body-create-setvalue-drops-run-properties/build.mjs
// Body paragraph 30000001: "Retention: " + "{{fancy_field}}" (styled run) + " of the sum."
// The token run carries direct formatting: Georgia, bold, italic, colour C00000, 12pt, language pt-BR.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
export const FANCY = '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:i/><w:color w:val="C00000"/><w:sz w:val="24"/><w:lang w:val="pt-BR"/></w:rPr>';
writeFixtures(join(HERE, 'fixtures'), [
  ['styled-token-body', await buildDocx({ body: p('30000001', r('Retention: ') + r('{{fancy_field}}', FANCY) + r(' of the sum.')) + p('30000002', r('Lorem ipsum.')) })],
]);
