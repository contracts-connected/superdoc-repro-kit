// Builds this case's fixture. Run: node cases/<this case>/build.mjs
// Six plain paragraphs with stable paraIds. No content control is hand-authored: the script mints every control
// through create.contentControl, so the controls have exactly the shape the host itself produces.
//   00000001 body       plain prose, outside every permission range
//   00000002 buyer_sig  -> block rich-text control (party 1's signature; the ONE permitted range)
//   00000003 buyer_text -> inline text control
//   00000004 buyer_date -> inline date control
//   00000005 buyer_chk  -> inline checkbox control
//   00000006 seller_sig -> block rich-text control (party 2's signature; NOT permitted)
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('00000001', r('Contract body text before the signature block.'))
  + p('00000002', r('buyer signature placeholder'))
  + p('00000003', r('buyer text placeholder'))
  + p('00000004', r('buyer date placeholder'))
  + p('00000005', r('buyer checkbox placeholder'))
  + p('00000006', r('seller signature placeholder'));
writeFixtures(join(HERE, 'fixtures'), [['signing-block', await buildDocx({ body })]]);
