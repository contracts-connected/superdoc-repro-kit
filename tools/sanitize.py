#!/usr/bin/env python3
"""Sanitize a real .docx template so it can be shared, keeping every structure a repro depends on.

What changes
  * Every word inside <w:t>, <w:delText> and VML textpath strings is replaced by a pseudo-word of the SAME
    length and case pattern (digits are remapped too), so character offsets, run boundaries, paraIds and
    the paragraph buffer the Document API computes stay identical.
  * Placeholder tokens are kept verbatim across run boundaries: {{code}}, [code], [code} and {code}.
  * A short list of generic contract words (Initials, Page, Buyer, Seller, ...) is kept for readability.
  * Tracked-change / comment authors -> "Author A", "Author B" ...; initials -> "AA".
  * docProps (title, creator, company ...), custom properties and customXml text are blanked.
  * External relationship targets and HYPERLINK field URLs -> https://example.com/
  * Embedded fonts (licensed binaries) are removed.
  * Images are replaced by a neutral placeholder of the same pixel size (EMF/WMF become PNG; the
    relationship target and content type are updated). The package thumbnail is removed.

What stays
  * All XML structure: paragraphs, runs, rPr/pPr, sectPr, tables, SDTs (tags, aliases, ids, locks),
    fields, bookmarks, numbering, styles, headers/footers, tracked changes, w14:paraId / w14:textId.

Usage: python tools/sanitize.py <in.docx> <out.docx>
Prints a leak check: any word of 4+ letters from the original text that still appears in the output.
"""
import hashlib, io, re, sys, zipfile
from html import unescape

try:
    from PIL import Image, ImageDraw
except ImportError:  # images are then replaced by a 1x1 PNG
    Image = None

KEEP = set("""
initials initial page of buyer seller signature signatures sign signed date dated name title by and the
purchase order no number contract contractor subcontractor subcontract agreement work scope amount total
project company address phone email terms conditions item items qty quantity unit price cost code description
schedule values value payment tax exhibit section article clause party parties owner vendor supplier
""".split())

TOKEN = re.compile(r"\{\{[^{}\n]{1,80}\}\}|\[[A-Za-z0-9_ .#:/-]{1,80}[\]\}]|\{[A-Za-z_][A-Za-z0-9_ .:-]{0,80}\}")
TEXT_NODE = re.compile(r"(<w:(?:t|delText)(?:\s[^>]*)?>)([^<]*)(</w:(?:t|delText)>)")
VOWELS, CONS = "aeiou", "bcdfghjklmnprstvw"


def pseudo(word):
    """Deterministic pseudo-word with the same length and per-character case pattern."""
    if word.lower() in KEEP:
        return word
    h = hashlib.sha256(word.lower().encode()).digest()
    out = []
    for i, ch in enumerate(word):
        pool = CONS if i % 2 == 0 else VOWELS
        c = pool[h[i % len(h)] % len(pool)]
        out.append(c.upper() if ch.isupper() else c)
    return "".join(out)


def scramble(text, protected):
    """Replace letters/digits outside protected positions; keeps length and punctuation."""
    out, i, n = list(text), 0, len(text)
    while i < n:
        if protected[i]:
            i += 1
            continue
        if text[i].isalpha():
            j = i
            while j < n and text[j].isalpha() and not protected[j]:
                j += 1
            out[i:j] = pseudo(text[i:j])
            i = j
        elif text[i].isdigit():
            out[i] = str((int(text[i]) * 7 + 3) % 10)
            i += 1
        else:
            i += 1
    return "".join(out)


def xml_escape(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def sanitize_text_nodes(xml):
    nodes = list(TEXT_NODE.finditer(xml))
    if not nodes:
        return xml
    texts = [unescape(m.group(2)) for m in nodes]
    joined = "".join(texts)
    protected = [False] * len(joined)
    for m in TOKEN.finditer(joined):
        for k in range(m.start(), m.end()):
            protected[k] = True
    scr = scramble(joined, protected)
    out, pos, last = [], 0, 0
    for m, t in zip(nodes, texts):
        new = scr[pos:pos + len(t)]
        pos += len(t)
        out.append(xml[last:m.start(2)])
        out.append(xml_escape(new))
        last = m.end(2)
    out.append(xml[last:])
    return "".join(out)


AUTHORS = {}


def author_alias(name):
    if name not in AUTHORS:
        AUTHORS[name] = "Author " + chr(ord("A") + len(AUTHORS) % 26)
    return AUTHORS[name]


def sanitize_word_xml(xml):
    xml = sanitize_text_nodes(xml)
    xml = re.sub(r'(w:author=")([^"]*)(")', lambda m: m.group(1) + author_alias(m.group(2)) + m.group(3), xml)
    xml = re.sub(r'(w:initials=")([^"]*)(")', r'\1AA\3', xml)
    xml = re.sub(r'(<wp:docPr\b[^>]*?\bname=")([^"]*)(")', r'\1Image\3', xml)
    xml = re.sub(r'(<wp:docPr\b[^>]*?\b(?:descr|title)=")([^"]*)(")', r'\1\3', xml)
    xml = re.sub(r'(<pic:cNvPr\b[^>]*?\bname=")([^"]*)(")', r'\1Image\3', xml)
    xml = re.sub(r'(<pic:cNvPr\b[^>]*?\bdescr=")([^"]*)(")', r'\1\3', xml)
    xml = re.sub(r'(HYPERLINK\s+(?:&quot;|"))([^"&]*)', r'\1https://example.com/', xml)
    xml = re.sub(r'(<v:textpath\b[^>]*?\bstring=")([^"]*)(")', lambda m: m.group(1) + xml_escape(scramble(unescape(m.group(2)), [False] * len(unescape(m.group(2))))) + m.group(3), xml)
    xml = re.sub(r'(<w:docVar\b[^>]*?\bw:val=")([^"]*)(")', r'\1\3', xml)
    # people.xml
    xml = re.sub(r'(<w15:person\b[^>]*?\bw15:author=")([^"]*)(")', lambda m: m.group(1) + author_alias(m.group(2)) + m.group(3), xml)
    xml = re.sub(r'(w15:userId=")([^"]*)(")', r'\1sanitized\3', xml)
    return xml


def blank_text(xml):
    # text nodes only: the captured text may contain neither '<' nor '>' so a match never spans a tag
    return re.sub(r">([^<>]*[^<>\s][^<>]*)<", ">sanitized<", xml)


def placeholder_image(data, fmt):
    if Image is None:
        return None
    try:
        im = Image.open(io.BytesIO(data))
        w, h = im.size
    except Exception:
        w, h = 200, 80
    w, h = max(1, min(w, 4000)), max(1, min(h, 4000))
    ph = Image.new("RGB", (w, h), (214, 219, 224))
    d = ImageDraw.Draw(ph)
    d.rectangle([0, 0, w - 1, h - 1], outline=(150, 158, 166))
    d.line([0, 0, w - 1, h - 1], fill=(150, 158, 166))
    d.line([0, h - 1, w - 1, 0], fill=(150, 158, 166))
    buf = io.BytesIO()
    ph.save(buf, format=fmt)
    return buf.getvalue()


RASTER = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG", ".gif": "GIF", ".bmp": "BMP", ".tif": "TIFF", ".tiff": "TIFF"}
VECTOR = {".emf", ".wmf"}


def main(src, dst):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    originals_text = []
    renamed = {}  # old media path -> new media path
    out_entries = []
    for name in names:
        data = zin.read(name)
        low = name.lower()
        if low.startswith("docprops/thumbnail") or low.startswith("word/fonts/"):
            continue  # thumbnail shows the original page; embedded fonts are licensed binaries
        if low.startswith("word/") and low.endswith(".xml"):
            xml = data.decode("utf-8")
            originals_text.append(" ".join(unescape(m.group(2)) for m in TEXT_NODE.finditer(xml)))
            xml = sanitize_word_xml(xml)
            if low == "word/fonttable.xml":
                xml = re.sub(r"<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*/>", "", xml)
            data = xml.encode("utf-8")
        elif low == "docprops/core.xml":
            xml = data.decode("utf-8")
            originals_text.append(re.sub(r"<[^>]+>", " ", xml))
            xml = re.sub(r"(<(dc:title|dc:subject|dc:creator|cp:keywords|dc:description|cp:lastModifiedBy|cp:category|cp:contentStatus)(?:\s[^>]*)?>)([^<]*)(</\2>)", r"\1Sanitized\4", xml)
            data = xml.encode("utf-8")
        elif low == "docprops/app.xml":
            xml = data.decode("utf-8")
            originals_text.append(re.sub(r"<[^>]+>", " ", xml))
            xml = re.sub(r"(<((?:\w+:)?(?:Company|Manager|HyperlinkBase|Template|Application|AppVersion))>)([^<]*)(</\2>)", r"\1\4", xml)
            xml = re.sub(r"(<vt:lpstr>)([^<]*)(</vt:lpstr>)", r"\1Sanitized\3", xml)
            data = xml.encode("utf-8")
        elif low in ("docprops/custom.xml",) or low.startswith("customxml/item") and low.endswith(".xml") and "props" not in low:
            xml = data.decode("utf-8", "replace")
            originals_text.append(re.sub(r"<[^>]+>", " ", xml))
            data = blank_text(xml).encode("utf-8")
        elif low.endswith(".rels"):
            xml = data.decode("utf-8")
            xml = re.sub(r'(<Relationship\b[^>]*?Target=")([^"]*)("[^>]*?TargetMode="External")', r"\1https://example.com/\3", xml)
            xml = re.sub(r'(<Relationship\b[^>]*?TargetMode="External"[^>]*?Target=")([^"]*)(")', r"\1https://example.com/\3", xml)
            if low == "word/_rels/fonttable.xml.rels":
                xml = re.sub(r'<Relationship\b[^>]*?Target="fonts/[^"]*"[^>]*/>', "", xml)
            if name == "_rels/.rels":
                xml = re.sub(r'<Relationship\b[^>]*?Target="docProps/thumbnail[^"]*"[^>]*/>', "", xml)
            data = xml.encode("utf-8")
        elif low.startswith("word/media/"):
            ext = "." + low.rsplit(".", 1)[-1]
            if ext in RASTER:
                new = placeholder_image(data, RASTER[ext])
                if new:
                    data = new
            elif ext in VECTOR:
                new = placeholder_image(b"", "PNG")
                if new:
                    newname = name.rsplit(".", 1)[0] + ".png"
                    renamed[name] = newname
                    name, data = newname, new
        out_entries.append((name, data))

    # rewrite relationship targets for renamed media, and make sure png has a content type
    if renamed:
        fixed = []
        for name, data in out_entries:
            if name.endswith(".rels"):
                xml = data.decode("utf-8")
                for old, new in renamed.items():
                    xml = xml.replace(old.split("/")[-1], new.split("/")[-1])
                data = xml.encode("utf-8")
            elif name == "[Content_Types].xml":
                xml = data.decode("utf-8")
                if 'Extension="png"' not in xml:
                    xml = xml.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1)
                for old in renamed:
                    xml = re.sub(r'<Override PartName="/' + re.escape(old) + r'"[^>]*/>', "", xml)
                data = xml.encode("utf-8")
            fixed.append((name, data))
        out_entries = fixed
    if any(n.lower().startswith("docprops/thumbnail") for n in names):
        out_entries = [(n, (d.decode("utf-8").replace('<Override PartName="/docProps/thumbnail.jpeg" ContentType="image/jpeg"/>', "").encode("utf-8") if n == "[Content_Types].xml" else d)) for n, d in out_entries]

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in out_entries:
            zi = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)

    # leak check: every original word of 4+ letters (not generic, not inside a token) must be gone
    zchk = zipfile.ZipFile(dst)
    blob = []
    for n in zchk.namelist():
        if n.endswith(".xml") or n.endswith(".rels"):
            x = zchk.read(n).decode("utf-8", "replace")
            blob.append(" ".join(unescape(m.group(2)) for m in TEXT_NODE.finditer(x)))
            blob.append(re.sub(r"<[^>]+>", " ", x) if n.startswith("docProps") or n.startswith("customXml") else "")
    after = " ".join(blob)
    orig = " ".join(originals_text)
    orig_wo_tokens = TOKEN.sub(" ", orig)
    words = {w for w in re.findall(r"[A-Za-z][A-Za-z]{3,}", orig_wo_tokens) if w.lower() not in KEEP}
    after_wo_tokens = TOKEN.sub(" ", after)
    after_words = set(re.findall(r"[A-Za-z][A-Za-z]{3,}", after_wo_tokens))
    leaks = sorted(w for w in words if w in after_words)
    kept_tokens = sorted(set(TOKEN.findall(after)))
    print(f"{src} -> {dst}")
    print(f"  parts: {len(out_entries)}  media renamed: {len(renamed)}  authors aliased: {len(AUTHORS)}")
    print(f"  tokens kept ({len(kept_tokens)}): {', '.join(kept_tokens[:25])}{' ...' if len(kept_tokens) > 25 else ''}")
    print(f"  LEAK CHECK: {'OK, no original word left' if not leaks else 'LEAKS: ' + ', '.join(leaks[:50])}")
    return 0 if not leaks else 3


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
