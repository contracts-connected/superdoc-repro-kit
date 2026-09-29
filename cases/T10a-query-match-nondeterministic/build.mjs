// Builds this case's fixtures. Run: node cases/T10a-query-match-nondeterministic/build.mjs
// One legacy bracket placeholder split across two runs, the first run bold.
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildDocx, writeFixtures, p, r } from '../../lib/docx.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const body = p('00000001', r('Buyer: ') + r('[Buying Company', '<w:rPr><w:b/></w:rPr>') + r(' Name] agrees to the terms.')) + p('00000002', r('Lorem ipsum dolor sit amet.'));
// A 1.0-style loop marker inside a table cell (the shape where omitting `in` returned 0).
const cell = (inner) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${cell(p('00000011', r('{{loop scope_items}}')))}${cell(p('00000012', r('Amount')))}</w:tr></w:tbl>`;
writeFixtures(join(HERE, 'fixtures'), [
  ['bracket-split-across-runs', await buildDocx({ body })],
  ['loop-marker-in-table-cell', await buildDocx({ body: p('00000001', r('Scope items:')) + table + p('00000002', r('Lorem ipsum.')) })],
]);
