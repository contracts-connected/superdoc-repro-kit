// Shared harness for every case: pick an SDK version, open a fixture, save, read XML parts, print a verdict.
//
// Every case is run the same way from the repo root:
//   node cases/<case-id>/repro.mjs [sdkVersion]        (default 2.13.0; must be a folder under ./sdk/)
//
// Exit codes (read by run-all.mjs):
//   0  BUG REPRODUCED   the behaviour described in the case README happened
//   1  FIXED?           the behaviour we expect happened instead
//   2  ERROR            the script could not complete (setup problem, crash) - never counts as fixed
import { readFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import JSZip from 'jszip';

export { findSdt, visible, rPrOf, IMAGE_SRC } from './xml.mjs';

export const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
export const CASE_DIR = dirname(process.argv[1]);
export const CASE_ID = basename(CASE_DIR);

for (const ev of ['uncaughtException', 'unhandledRejection']) {
  process.on(ev, (e) => {
    console.log(`ERROR (script did not complete): ${e?.code ?? ''} ${e?.message ?? e}`);
    if (e?.details) console.log(JSON.stringify(e.details));
    console.log(`RESULT ${JSON.stringify({ case: CASE_ID, sdk: SDK_VERSION, verdict: 'ERROR', error: String(e?.message ?? e) })}`);
    process.exit(2);
  });
}

export const SDK_VERSION = process.argv[2] || process.env.SUPERDOC_SDK || '2.13.0';
export const sha256 = (b) => createHash('sha256').update(b).digest('hex');
export const short = (x, n = 300) => { const s = typeof x === 'string' ? x : JSON.stringify(x); return s && s.length > n ? s.slice(0, n) + '...' : s; };

const sdkDir = join(ROOT, 'sdk', SDK_VERSION);
if (!existsSync(join(sdkDir, 'node_modules/@superdoc/sdk/package.json'))) {
  console.error(`@superdoc/sdk ${SDK_VERSION} is not installed. Run: npm ci --prefix sdk/${SDK_VERSION}`);
  console.log(`RESULT ${JSON.stringify({ case: CASE_ID, sdk: SDK_VERSION, verdict: 'ERROR', error: 'sdk not installed' })}`);
  process.exit(2);
}
const sdkPkg = JSON.parse(await readFile(join(sdkDir, 'node_modules/@superdoc/sdk/package.json'), 'utf8'));
const { SuperDocClient } = await import(pathToFileURL(join(sdkDir, 'node_modules/@superdoc/sdk/dist/index.js')).href);
export { SuperDocClient };

// ---- calling the SDK ----------------------------------------------------------------------------
// A throw becomes { THROW: 'CODE message' } so the script can print it and keep going.
export const attempt = async (fn) => { try { return await fn(); } catch (e) { return { THROW: `${e?.code ?? ''} ${e?.message ?? e}`.trim() }; } };
export const ok = (r) => r != null && !r.THROW && r.success !== false;
export const describe = (r) => (r?.THROW ? `THROWS ${r.THROW}` : r?.success === false ? `success:false ${short(r.failure ?? r, 200)}` : `success:${r?.success ?? '(value returned)'}`);
export const hasFn = (obj, path) => { let o = obj; for (const k of path.split('.')) { if (o == null) return false; o = o[k]; } return typeof o === 'function'; };

// ---- addressing shorthands (the exact shapes the Document API documents) --------------------------
export const BODY = { kind: 'story', storyType: 'body' };
export const hfStory = (refId) => ({ kind: 'story', storyType: 'headerFooterPart', refId });
export const pt = (blockId, offset, story) => (story ? { kind: 'text', blockId, offset, story } : { kind: 'text', blockId, offset });
export const sel = (start, end) => ({ kind: 'selection', start, end });
export const range = (blockId, s, e, story) => sel(pt(blockId, s, story), pt(blockId, e, story));
export const sdtBlock = (nodeId) => ({ kind: 'block', nodeType: 'sdt', nodeId });
export const sdtInline = (nodeId) => ({ kind: 'inline', nodeType: 'sdt', nodeId });
export const paragraph = (nodeId, story) => (story ? { kind: 'block', nodeType: 'paragraph', nodeId, story } : { kind: 'block', nodeType: 'paragraph', nodeId });

// ---- fixtures ------------------------------------------------------------------------------------
// Looked up in: the case's own fixtures/ folder, then fixtures/sanitized/, then fixtures/synthetic/.
export function fixturePath(name) {
  const file = name.endsWith('.docx') ? name : `${name}.docx`;
  for (const dir of [join(CASE_DIR, 'fixtures'), join(ROOT, 'fixtures', 'sanitized'), join(ROOT, 'fixtures', 'synthetic')]) {
    if (existsSync(join(dir, file))) return join(dir, file);
  }
  throw new Error(`fixture not found: ${file}`);
}

export async function header(title, fixtures = []) {
  console.log('='.repeat(100));
  console.log(`${CASE_ID}: ${title}`);
  console.log(`@superdoc/sdk ${sdkPkg.version} | node ${process.version} | ${process.platform}-${process.arch} | ${new Date().toISOString()}`);
  for (const f of fixtures) console.log(`fixture ${f}${f.endsWith('.docx') ? '' : '.docx'} sha256=${sha256(await readFile(fixturePath(f)))}`);
  console.log('-'.repeat(100));
}

// Open a fixture, run fn(doc), optionally save to out/<case>/<outName>-sdk-<v>.docx, close.
// openOptions are passed to client.open (e.g. { userName, userEmail }). Returns { value, bytes }.
export async function withDoc(fixture, fn, outName = null, openOptions = {}) {
  const client = new SuperDocClient();
  await client.connect();
  try {
    const doc = await client.open({ doc: fixturePath(fixture), ...openOptions });
    let value, bytes = null;
    try {
      value = await fn(doc);
      if (outName) {
        const dir = join(ROOT, 'out', CASE_ID);
        await mkdir(dir, { recursive: true });
        const out = join(dir, `${outName}-sdk-${SDK_VERSION}.docx`);
        await doc.save({ out, force: true });
        bytes = await readFile(out);
      }
    } finally {
      await doc.close({ discard: true }).catch(() => {});
    }
    return { value, bytes };
  } finally {
    await client.dispose();
  }
}

// Re-open saved bytes (round-trip checks). Returns fn's value.
export async function reopen(bytes, fn) {
  const dir = join(ROOT, 'out', CASE_ID);
  await mkdir(dir, { recursive: true });
  const p = join(dir, `reopen-${Date.now()}-sdk-${SDK_VERSION}.docx`);
  await (await import('node:fs/promises')).writeFile(p, bytes);
  const client = new SuperDocClient();
  await client.connect();
  try {
    const doc = await client.open({ doc: p });
    try { return await fn(doc); } finally { await doc.close({ discard: true }).catch(() => {}); }
  } finally { await client.dispose(); }
}

// ---- package inspection --------------------------------------------------------------------------
export async function part(bytesOrPath, name) {
  const bytes = typeof bytesOrPath === 'string' ? await readFile(bytesOrPath) : bytesOrPath;
  const z = await JSZip.loadAsync(bytes);
  return z.file(name) ? z.file(name).async('string') : null;
}
export async function partNames(bytesOrPath) {
  const bytes = typeof bytesOrPath === 'string' ? await readFile(bytesOrPath) : bytesOrPath;
  return Object.keys((await JSZip.loadAsync(bytes)).files);
}

// ---- verdict -------------------------------------------------------------------------------------
// reproduced=true -> exit 0 "BUG REPRODUCED"; false -> exit 1 "FIXED?".
export function verdict(reproduced, expected, actual, extra = {}) {
  console.log('-'.repeat(100));
  console.log(`EXPECTED: ${expected}`);
  console.log(`ACTUAL:   ${actual}`);
  console.log(reproduced ? `BUG REPRODUCED (sdk ${SDK_VERSION})` : `FIXED? behaviour is correct on sdk ${SDK_VERSION}`);
  console.log(`RESULT ${JSON.stringify({ case: CASE_ID, sdk: sdkPkg.version, verdict: reproduced ? 'REPRODUCED' : 'FIXED', actual, ...extra })}`);
  process.exit(reproduced ? 0 : 1);
}
