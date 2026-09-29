// Deterministic builder for the synthetic fixtures. Every case that needs its own fixtures has a
// build.mjs next to repro.mjs that calls these helpers and writes cases/<case>/fixtures/*.docx.
// Fixed zip timestamps + fixed DEFLATE level: `node cases/<case>/build.mjs` reproduces the committed
// bytes exactly (with jszip 3.10.1 from the root package-lock.json). Synthetic text only.
import JSZip from 'jszip';
import { mkdirSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const DATE = new Date('2026-01-01T00:00:00Z');
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS = `xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`;
export const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// Run properties used across fixtures (font + bold + colour + size) - the rPr a mint must preserve.
export const RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="0B5394"/><w:sz w:val="20"/></w:rPr>';
export const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const r = (t, rpr = '') => `<w:r>${rpr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
export const tab = (rpr = '') => `<w:r>${rpr}<w:tab/></w:r>`;
export const br = (rpr = '') => `<w:r>${rpr}<w:br/></w:r>`;
export const p = (paraId, inner, pPr = '') => `<w:p w14:paraId="${paraId}" w14:textId="77777777">${pPr}${inner}</w:p>`;
export const pageField = (rpr = '') => `${r('Page ', rpr)}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
export const richSdt = (tag, id, txt, rpr = '', lock = '') => `<w:sdt><w:sdtPr>${rpr}<w:alias w:val="${tag}"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/>${lock ? `<w:lock w:val="${lock}"/>` : ''}</w:sdtPr><w:sdtContent><w:r>${rpr}<w:t>${esc(txt)}</w:t></w:r></w:sdtContent></w:sdt>`;
export const plainSdt = (tag, id, txt, rpr = '') => `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:text/></w:sdtPr><w:sdtContent><w:r>${rpr}<w:t>${esc(txt)}</w:t></w:r></w:sdtContent></w:sdt>`;
export const ftr = (inner) => `${XML}<w:ftr ${NS}>${inner}</w:ftr>`;
export const hdr = (inner) => `${XML}<w:hdr ${NS}>${inner}</w:hdr>`;
export const PG = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

const CT = {
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
};
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Build a minimal, valid .docx.
 * @param {object} o
 * @param {string}   o.body        inner XML of <w:body> WITHOUT the final <w:sectPr> (paragraphs, tables, inner sectPr in pPr)
 * @param {string[]} [o.footers]   footer part XML strings -> word/footer1.xml ... (relationship ids rIdF1 ...)
 * @param {string[]} [o.headers]   header part XML strings -> word/header1.xml ... (relationship ids rIdH1 ...)
 * @param {Array}    [o.footerRefs] [[type, index1Based]] for the final sectPr (default: [['default',1]] when footers exist)
 * @param {Array}    [o.headerRefs] [[type, index1Based]] for the final sectPr
 * @param {boolean}  [o.titlePg]
 * @param {string}   [o.numbering] word/numbering.xml content
 * @param {string}   [o.finalSectPr] full override of the final <w:sectPr>…</w:sectPr>
 * @returns {Promise<Buffer>}
 */
export async function buildDocx(o) {
  const footers = o.footers ?? [];
  const headers = o.headers ?? [];
  const z = new JSZip();
  const add = (n, c) => z.file(n, c, { date: DATE, createFolders: false });
  add('[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${footers.map((_, i) => `<Override PartName="/word/footer${i + 1}.xml" ContentType="${CT.footer}"/>`).join('')}${headers.map((_, i) => `<Override PartName="/word/header${i + 1}.xml" ContentType="${CT.header}"/>`).join('')}${o.numbering ? `<Override PartName="/word/numbering.xml" ContentType="${CT.numbering}"/>` : ''}</Types>`);
  add('_rels/.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${footers.map((_, i) => `<Relationship Id="rIdF${i + 1}" Type="${REL}/footer" Target="footer${i + 1}.xml"/>`).join('')}${headers.map((_, i) => `<Relationship Id="rIdH${i + 1}" Type="${REL}/header" Target="header${i + 1}.xml"/>`).join('')}${o.numbering ? `<Relationship Id="rIdN1" Type="${REL}/numbering" Target="numbering.xml"/>` : ''}</Relationships>`);
  footers.forEach((f, i) => add(`word/footer${i + 1}.xml`, f));
  headers.forEach((h, i) => add(`word/header${i + 1}.xml`, h));
  if (o.numbering) add('word/numbering.xml', o.numbering);
  const fRefs = o.footerRefs ?? (footers.length ? [['default', 1]] : []);
  const hRefs = o.headerRefs ?? [];
  const sect = o.finalSectPr ?? `<w:sectPr>${hRefs.map(([t, i]) => `<w:headerReference w:type="${t}" r:id="rIdH${i}"/>`).join('')}${fRefs.map(([t, i]) => `<w:footerReference w:type="${t}" r:id="rIdF${i}"/>`).join('')}${PG}${o.titlePg ? '<w:titlePg/>' : ''}</w:sectPr>`;
  add('word/document.xml', `${XML}<w:document ${NS}><w:body>${o.body}${sect}</w:body></w:document>`);
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

// Write fixtures to <dir>/<name>.docx and <dir>/SHA256SUMS. entries: [[name, Buffer]]
export function writeFixtures(dir, entries) {
  mkdirSync(dir, { recursive: true });
  const sums = [];
  for (const [name, bytes] of entries) {
    writeFileSync(join(dir, `${name}.docx`), bytes);
    sums.push(`${createHash('sha256').update(bytes).digest('hex')}  ${name}.docx`);
  }
  writeFileSync(join(dir, 'SHA256SUMS'), sums.join('\n') + '\n');
  console.log(sums.join('\n'));
}
