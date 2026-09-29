// Pure XML helpers shared by every case. No dependencies.
import zlib from 'node:zlib';

// The <w:sdt> element whose sdtPr has w:tag=tag (handles nested SDTs).
export function findSdt(xml, tag) {
  const tagAt = xml.indexOf(`<w:tag w:val="${tag}"/>`);
  if (tagAt < 0) return null;
  const start = xml.lastIndexOf('<w:sdt>', tagAt);
  let depth = 0, end = -1;
  for (const m of xml.slice(start).matchAll(/<w:sdt>|<\/w:sdt>/g)) {
    depth += m[0] === '<w:sdt>' ? 1 : -1;
    if (depth === 0) { end = start + m.index + m[0].length; break; }
  }
  const el = xml.slice(start, end);
  const pr = (el.match(/<w:sdtPr>([\s\S]*?)<\/w:sdtPr>/) || [])[1] || '';
  const cStart = el.indexOf('<w:sdtContent>') + '<w:sdtContent>'.length;
  const content = el.slice(cStart, el.lastIndexOf('</w:sdtContent>'));
  return { xml: el, index: start, pr, content };
}
// Human-readable text of an XML fragment: text, <TAB>, {DRAWING}, {FIELD ...}
export function visible(xml) {
  let t = '';
  for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:drawing>|<w:fldChar w:fldCharType="(\w+)"\/>|<w:sdtContent>|<\/w:sdtContent>/g)) {
    if (m[1] !== undefined) t += m[1];
    else if (m[0] === '<w:tab/>') t += '<TAB>';
    else if (m[0] === '<w:drawing>') t += '{DRAWING}';
    else if (m[0] === '<w:sdtContent>') t += '[[';
    else if (m[0] === '</w:sdtContent>') t += ']]';
    else t += `{fld:${m[2]}}`;
  }
  return t;
}
export const rPrOf = (xmlFragment) => (xmlFragment.match(/<w:r>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [])[1] ?? '(no rPr)';

// Deterministic 72x28 PNG (a blue squiggle) used as the "initials" image.
function png(w, h) {
  const tbl = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; tbl[n] = c >>> 0; }
  const crc = (b) => { let c = 0xffffffff; for (const x of b) c = tbl[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([l, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (w * 4 + 1) + 1 + x * 4; raw[o] = 0x10; raw[o + 1] = 0x20; raw[o + 2] = 0x80; raw[o + 3] = Math.abs(y - 14 + Math.round(6 * Math.sin(x / 6))) < 2 ? 255 : 0; }
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}
export const IMAGE_SRC = 'data:image/png;base64,' + png(72, 28).toString('base64');

