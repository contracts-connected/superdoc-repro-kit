// Builds this case's fixtures. Run: node cases/T01b-create-before-tab-deletes-tab/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const tail = p('10000009', r('Closing paragraph.'));
writeFixtures(join(HERE, 'fixtures'), [
  // ONE run holding two placeholders separated by a tab
  ['one-run-two-tokens-tab', await buildDocx({ body: p('00000006', `<w:r>${t('{{f1}}')}<w:tab/>${t('{{f2}}')}</w:r>`) + tail })],
  // ONE run: prose, tab, then the placeholder followed by more prose
  ['one-run-prose-tab-token', await buildDocx({ body: p('0000000D', `<w:r>${t('Line one prose.')}<w:tab/>${t('{{t}} more prose.')}</w:r>`) + tail })],
]);
