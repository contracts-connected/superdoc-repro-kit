// Builds this case's fixtures. Run: node cases/T10b-textbox-content-unreachable/build.mjs
// A body paragraph, then a paragraph holding an inline drawing whose text box (wps:txbx) contains text.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const textbox = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="TextBox 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="00000009" w14:textId="77777777"><w:r><w:t xml:space="preserve">Textbox content {{project_name}}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
writeFixtures(join(HERE, 'fixtures'), [
  ['textbox', await buildDocx({ body: p('00000001', r('Body text before the text box.')) + p('00000002', textbox) + p('00000003', r('Body text after the text box.')) })],
]);
