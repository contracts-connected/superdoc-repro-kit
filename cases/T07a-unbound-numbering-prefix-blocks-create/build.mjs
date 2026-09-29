// Builds this case's fixtures. Run: node cases/T07a-unbound-numbering-prefix-blocks-create/build.mjs
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, XML, W, ftr, p, r, RPR } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
// numbering.xml as SuperDoc 1.45.x exported it: w15:/w16cid: attributes, only xmlns:w declared on the root.
const numbering = (rootAttrs) => `${XML}<w:numbering ${rootAttrs}><w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1" w16cid:durableId="100000001"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const numPr = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const body = p('10000001', r('Lorem ipsum dolor sit amet.')) + p('10000002', r('Buyer: {{buyer_name}}')) + p('10000003', r('Lorem list item.'), numPr);
const footers = [ftr(p('50000001', r('Initials: [buyer_initials]', RPR)))];
writeFixtures(join(HERE, 'fixtures'), [
  ['numbering-undeclared-prefixes', await buildDocx({ body, footers, numbering: numbering(`xmlns:w="${W}"`) })],
  ['numbering-declared-prefixes', await buildDocx({ body, footers, numbering: numbering(`xmlns:w="${W}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"`) })],
]);
