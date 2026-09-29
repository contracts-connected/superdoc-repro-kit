// Builds this case's fixtures. Run: node cases/T01a-create-after-break-deletes-siblings/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const tail = p('10000009', r('Closing paragraph.'));
writeFixtures(join(HERE, 'fixtures'), [
  // ONE run holding two placeholders separated by one hard line break
  ['one-run-two-tokens-br', await buildDocx({ body: p('00000005', `<w:r>${t('{{f1}}')}<w:br/>${t('{{f2}}')}</w:r>`) + tail })],
  // ONE run: prose, hard line break, then the placeholder followed by more prose
  ['one-run-prose-br-token', await buildDocx({ body: p('0000000A', `<w:r>${t('Line one prose.')}<w:br/>${t('{{t}} more prose.')}</w:r>`) + tail })],
]);
