"""Rule-based HTML cleaner: the H&D_improved_prompt.md rules as Python code.

Input per product: Name, Highlights (English/Bengali), Description (English/Bengali).
Output, matching Claude's real output format:
  Highlights:  <ul><li>..</li>..</ul>
  Description: <h2>title</h2> <ul>bullets</ul> <p>text</p> [<h3>..</h3>] spec lists/tables  <img src=".."/>

Main rules (see the prompt for the full wording):
  - only h2/h3/ul/li/p/table/tr/td/img are output; every other tag is unwrapped
    and every attribute except img src is removed
  - corrupted markup (leaked ' style="...">' text), Hebrew, emojis, broken symbols removed
  - "empty" = fewer than 3 readable characters
  - Highlights: both empty -> from the Description bullets (same language);
    one empty -> copy the other; paragraphs / '*' '•' '$' text -> bullets
  - Description: both empty -> built from the Highlights; one empty -> copy the other
  - junk bullets (no word of 3+ letters) dropped, near-duplicate bullets merged
  - a paragraph mostly repeating the bullets is dropped
  - every image of the row is kept, at the bottom of the descriptions
  - generic Highlights reused by 2+ products are replaced by the product's own
    Description bullets (when it has 4+ specific ones)
"""

import difflib
import html
import re
import unicodedata
from collections import Counter
from html.parser import HTMLParser

# ---------------------------------------------------------------------------
# Text cleaning
# ---------------------------------------------------------------------------
_LEAKED_ATTR = re.compile(r'\b(?:style|class|id|align|width|height|dir|lang|data-[\w-]+)\s*=\s*'
                          r'(?:"[^"]*"|\'[^\']*\')\s*/?>?', re.IGNORECASE)
_REMOVE_CHARS = re.compile(
    "[֐-׿"          # Hebrew
    "�￼"            # broken-character symbols
    "​-‍⁠﻿"  # zero-width
    "︎️"            # emoji variation selectors
    "\U0001F000-\U0001FAFF"   # emojis, pictographs
    "☀-➿"           # misc symbols / dingbats (✔ ★ ☎ ➤ ...)
    "⬀-⯿"           # arrows / stars
    "\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")
_BULLET_MARKS = "•●◦▪■□◆◇►▶➤✓✔✅❖☛→»"
_WS = re.compile(r"\s+")


def clean_text(text):
    text = _LEAKED_ATTR.sub(" ", text)
    text = _REMOVE_CHARS.sub(" ", text)
    # Stray < > from broken tags go; real comparisons like "< 20 mA" or ">= 5" stay.
    text = re.sub(r"<(?!\s*[\d=])", " ", text)
    text = re.sub(r"(?<![\d\s<-])>(?!\s*[\d=])|(?<=\s)>(?!\s*[\d=])", " ", text)
    return _WS.sub(" ", text).strip()


def real_len(text):
    """Readable characters (letters/digits) - used for the 'empty' rule."""
    return sum(1 for c in text if c.isalnum())


def _is_letter(c):
    # Bangla vowel signs (া ী ...) are marks, not letters, but belong to the word.
    return c.isalpha() or unicodedata.category(c) in ("Mn", "Mc")


def has_real_word(text):
    """At least one word with 3+ letters (works for Bangla too)."""
    return any(sum(1 for c in w if _is_letter(c)) >= 3 for w in text.split())


def is_junk_bullet(text):
    """A stray character / 1-2 letter fragment. '200 g' or '00000007658' are real (they have numbers)."""
    return real_len(text) < 3 or not (has_real_word(text) or any(c.isdigit() for c in text))


_STOP = set("""a an the and or of to in on for with by at from as is are was were be been this that these
those it its your you our we they their will can may has have had not no but if so than then into
up out over about per all any each more most very also just only such use used using""".split())


def content_words(text):
    words = ("".join(c if _is_letter(c) else " " for c in text.lower())).split()
    return {w for w in words if len(w) >= 3 and w not in _STOP}


def split_plain_bullets(text):
    """'* one * two', '• a • b', 'details$ Type: X$ Brand: Y' -> list of items (else [text])."""
    for mark in _BULLET_MARKS:
        if text.count(mark) >= 1 and len(text.split(mark)) >= 3 or text.startswith(mark):
            parts = [p.strip() for p in text.split(mark)]
            return [p for p in parts if p]
    star = r"(?:^|\s)\*\s+(?=[^\d\s*])"  # '* item' with spaces; not 54 * 45 or L*W*H
    if len(re.findall(star, text)) >= 2:
        return [p.strip() for p in re.split(star, text) if p.strip()]
    if text.count("$") >= 2 and not re.search(r"\$\s*\d", text):
        return [p.strip() for p in text.split("$") if p.strip()]
    return [text]


def split_sentences(text):
    parts = re.split(r"(?<=[.!?।])\s+(?=[A-Zঀ-৿])", text)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# HTML -> blocks
# ---------------------------------------------------------------------------
class _Parser(HTMLParser):
    """Turns messy HTML into blocks: ('heading', text) ('list', [items])
    ('para', text) ('table', [[cells]]) plus a list of image URLs."""

    BLOCK = {"p", "div", "section", "article", "blockquote", "header", "footer", "pre", "center"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.blocks, self.images = [], []
        self.buf = []
        self.heading = None
        self.list_depth = 0
        self.items, self.li = [], None
        self.table, self.row, self.cell = None, None, None

    # -- buffers
    def _flush_para(self):
        text = clean_text(" ".join(self.buf))
        self.buf = []
        if text:
            self.blocks.append(("para", text))

    def _flush_li(self):
        if self.li is not None:
            text = clean_text(" ".join(self.li))
            if text:
                self.items.append(text)
        self.li = None

    def _flush_list(self):
        self._flush_li()
        if self.items:
            self.blocks.append(("list", self.items))
        self.items = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "img":
            src = dict(attrs).get("src")
            if src and src.strip() and src.strip() not in self.images:
                self.images.append(src.strip())
        elif tag == "br":
            if self.cell is not None:
                self.cell.append(" ")
            elif self.li is not None:
                self.li.append(" ")
            elif self.list_depth == 0 and self.heading is None:
                self._flush_para()
        elif tag in ("ul", "ol"):
            if self.list_depth == 0:
                self._flush_para()
            self._flush_li()
            self.list_depth += 1
        elif tag == "li":
            if self.list_depth == 0:
                self._flush_para()
                self.list_depth = 1
            self._flush_li()
            self.li = []
        elif re.fullmatch(r"h[1-6]", tag) and self.cell is None and self.li is None:
            self._flush_para()
            self.heading = []
        elif tag == "table":
            self._flush_para()
            self._flush_list()
            self.list_depth = 0
            self.table = []
        elif tag == "tr" and self.table is not None:
            self.row = []
            self.table.append(self.row)
        elif tag in ("td", "th") and self.table is not None:
            if self.row is None:
                self.row = []
                self.table.append(self.row)
            self.cell = []
            self.row.append(self.cell)
        elif tag in self.BLOCK and self.list_depth == 0 and self.cell is None and self.heading is None:
            self._flush_para()
        elif tag in self.BLOCK and self.li is not None:
            self.li.append(" ")

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("ul", "ol") and self.list_depth:
            self._flush_li()
            self.list_depth -= 1
            if self.list_depth == 0:
                self._flush_list()
        elif tag == "li":
            self._flush_li()
        elif re.fullmatch(r"h[1-6]", tag) and self.heading is not None:
            text = clean_text(" ".join(self.heading))
            self.heading = None
            if text:
                self.blocks.append(("heading", text))
        elif tag in ("td", "th"):
            self.cell = None
        elif tag == "tr":
            self.row = None
        elif tag == "table" and self.table is not None:
            rows = [[clean_text(" ".join(c)) for c in r] for r in self.table]
            rows = [[c for c in r if real_len(c) > 0] for r in rows]  # no empty <td>
            rows = [r for r in rows if r]
            if rows:
                self.blocks.append(("table", rows))
            self.table, self.row, self.cell = None, None, None
        elif tag in self.BLOCK and self.list_depth == 0 and self.cell is None and self.heading is None:
            self._flush_para()

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)
        elif self.heading is not None:
            self.heading.append(data)
        elif self.li is not None:
            self.li.append(data)
        elif self.list_depth:
            self.li = [data]  # text directly inside <ul> without <li>
        else:
            self.buf.append(data)

    def finish(self):
        self.close()
        if self.heading is not None:
            text = clean_text(" ".join(self.heading))
            if text:
                self.blocks.append(("heading", text))
        self._flush_para()
        self._flush_list()
        return self.blocks, self.images


def parse(raw):
    p = _Parser()
    text = str(raw or "")
    # Repair markup whose '<tag' was lost: drop leaked attribute text before parsing.
    text = re.sub(r'(^|>)([^<]*?)\s(?:style|class)\s*=\s*"[^"]*"\s*/?>', r"\1\2 ", text)
    try:
        p.feed(text)
    except Exception:
        p = _Parser()
        p.feed(html.escape(re.sub(r"<[^>]*>", " ", text)))
    return p.finish()


# ---------------------------------------------------------------------------
# Bullet helpers
# ---------------------------------------------------------------------------
def tidy_bullets(items):
    """Split plain-text bullets, drop junk, remove exact/near duplicates and concatenations."""
    out = []
    for item in items:
        for part in split_plain_bullets(item):
            part = clean_text(part.lstrip(_BULLET_MARKS + "-–*$ "))
            if part and not is_junk_bullet(part):
                out.append(part)
    keys = [re.sub(r"[^\w]+", " ", b.lower()).strip() for b in out]
    words = [k.split() for k in keys]
    keep = [True] * len(out)
    for i in range(len(out)):
        if not keep[i]:
            continue
        for j in range(len(out)):
            if i == j or not keep[j] or not keep[i]:
                continue
            a, b = keys[i], keys[j]
            if a == b:
                keep[max(i, j)] = False
            elif len(a) <= len(b) and a in b and len(words[i]) >= 4 and len(a) >= 0.7 * len(b):
                keep[i] = False
            elif min(len(a), len(b)) > 20 and _similar(words[i], words[j]):
                keep[i if len(a) < len(b) else j] = False
    # A bullet that is just several other bullets glued together.
    for i in range(len(out)):
        if not keep[i]:
            continue
        inside = [keys[j] for j in range(len(out)) if j != i and keep[j] and keys[j] and keys[j] in keys[i]]
        if len(inside) >= 2 and sum(len(k) for k in inside) >= 0.8 * len(keys[i]):
            keep[i] = False
    return [b for b, k in zip(out, keep) if k]


def _similar(a, b):
    """Same bullet with a word or two missing/changed (word-level, quick checks first)."""
    if min(len(a), len(b)) < 0.8 * max(len(a), len(b)):
        return False
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    return sm.quick_ratio() >= 0.9 and sm.ratio() >= 0.9


def to_bullets(blocks):
    """Highlights content -> bullet list: list items, paragraphs, headings, tables."""
    items = []
    for kind, value in blocks:
        if kind == "list":
            items += value
        elif kind in ("para", "heading"):
            parts = split_plain_bullets(value)
            if len(parts) == 1:
                parts = split_sentences(value) if len(value) > 160 else [value]
            items += parts
        elif kind == "table":
            items += [": ".join(c for c in row if c) for row in value]
    return tidy_bullets(items)


def ul(items):
    return "<ul>" + "".join(f"<li>{html.escape(i, quote=False)}</li>" for i in items) + "</ul>" if items else ""


def has_text(blocks):
    return real_len(" ".join(_block_text(b) for b in blocks)) >= 3


def _block_text(block):
    kind, value = block
    if kind == "list":
        return " ".join(value)
    if kind == "table":
        return " ".join(" ".join(r) for r in value)
    return value


# ---------------------------------------------------------------------------
# Description
# ---------------------------------------------------------------------------
# A paragraph is "repeating the bullets" when this share of its meaningful words is in them.
REPEAT_SHARE = 0.85

_SPEC_HEADING = re.compile(r"spec|feature|detail|package|include|content|dimension|technical|"
                           r"বৈশিষ্ট্য|স্পেসিফিকেশন", re.IGNORECASE)


class Description:
    def __init__(self, blocks, images, name):
        self.images = list(images)
        self.title = None
        self.source_title = False
        self.bullets, self.paras, self.spec = [], [], []  # spec: [("h3"|"list"|"table", value)]
        pending_heading = None
        for i, (kind, value) in enumerate(blocks):
            if kind == "heading":
                if self.title is None and not self.bullets and not self.paras:
                    self.title = value
                    self.source_title = real_len(value) >= 3
                else:
                    pending_heading = value
                continue
            if kind == "list":
                is_spec = bool(self.bullets)  # a 2nd list goes to the spec section
                if not is_spec and not self.bullets:
                    self.bullets = list(value)
                elif not is_spec:
                    self.bullets += value
                else:
                    if pending_heading:
                        self.spec.append(("h3", pending_heading))
                    self.spec.append(("list", list(value)))
            elif kind == "table":
                if pending_heading and _SPEC_HEADING.search(pending_heading):
                    self.spec.append(("h3", pending_heading))
                self.spec.append(("table", value))
            elif kind == "para":
                if pending_heading and not _SPEC_HEADING.search(pending_heading) and \
                        len(pending_heading.split()) > 4 and pending_heading != self.title:
                    self.paras.append(pending_heading)
                parts = split_plain_bullets(value)
                if len(parts) > 1:
                    target = self.bullets if not self.spec else None
                    if target is not None:
                        self.bullets += parts
                    else:
                        self.spec.append(("list", parts))
                else:
                    self.paras.append(value)
            pending_heading = None
        self.bullets = tidy_bullets(self.bullets)
        self.spec = [(k, tidy_bullets(v) if k == "list" else v) for k, v in self.spec]
        self.spec = [(k, v) for k, v in self.spec if v]
        if not self.title:
            self.title = clean_text(str(name or ""))
        self.paras = [p for p in self.paras if real_len(p) >= 3]

    def has_body(self):
        return bool(self.bullets or self.paras or self.spec)

    def has_text(self):
        """'Empty' in the prompt's sense: no readable text at all (a real title counts)."""
        return self.source_title or self.has_body()

    def drop_repeated_paragraphs(self, bullets):
        words = set()
        for b in bullets:
            words |= content_words(b)
        for k, v in self.spec:
            if k == "list":
                for b in v:
                    words |= content_words(b)
        kept = []
        for p in self.paras:
            pw = content_words(p)
            if pw and len(pw & words) / len(pw) >= REPEAT_SHARE:
                continue
            kept.append(p)
        self.paras = kept

    def render(self):
        out = [f"<h2>{html.escape(self.title, quote=False)}</h2>"] if self.title else []
        out.append(ul(self.bullets))
        out += [f"<p>{html.escape(p, quote=False)}</p>" for p in self.paras]
        for kind, value in self.spec:
            if kind == "h3":
                out.append(f"<h3>{html.escape(value, quote=False)}</h3>")
            elif kind == "list":
                out.append(ul(value))
            else:
                out.append("<table>" + "".join(
                    "<tr>" + "".join(f"<td>{html.escape(c, quote=False)}</td>" for c in r) + "</tr>"
                    for r in value) + "</table>")
        out += [f'<img src="{html.escape(u)}"/>' for u in self.images]
        return "".join(out)


# ---------------------------------------------------------------------------
# One product
# ---------------------------------------------------------------------------
def clean_product(name, he, hb, de, db):
    """Returns dict with keys he, hb, de, db (clean HTML) and notes (what was derived)."""
    p_he, p_hb, p_de, p_db = parse(he), parse(hb), parse(de), parse(db)
    notes = set()

    H = {"E": to_bullets(p_he[0]), "B": to_bullets(p_hb[0])}
    D = {"E": Description(*p_de, name), "B": Description(*p_db, name)}
    own_desc_bullets = {k: list(D[k].bullets) for k in D}

    # Highlights
    if not H["E"] and not H["B"]:
        for k in "EB":
            H[k] = D[k].bullets or tidy_bullets(sum((split_sentences(p) for p in D[k].paras), []))[:8]
        if H["E"] or H["B"]:
            notes.add("highlights_from_description")
    for a, b in (("E", "B"), ("B", "E")):
        if not H[a] and H[b]:
            H[a] = list(H[b])
    if not H["E"] and not H["B"]:
        H["E"] = H["B"] = [clean_text(str(name or ""))] if name else []

    # Descriptions
    if not D["E"].has_text() and not D["B"].has_text():
        for k in "EB":
            D[k].bullets = list(H[k])
        notes.add("description_from_highlights")
    for a, b in (("E", "B"), ("B", "E")):
        if not D[a].has_text() and D[b].has_text():
            src = D[b]
            D[a].bullets, D[a].paras, D[a].spec = list(src.bullets), list(src.paras), list(src.spec)
            D[a].title, D[a].source_title = src.title, src.source_title

    # Images: every image of the row stays (own description first, then the rest).
    all_images = []
    for imgs in (p_de[1], p_he[1], p_db[1], p_hb[1]):
        for u in imgs:
            if u not in all_images:
                all_images.append(u)
    for k, own in (("E", p_de[1] + p_he[1]), ("B", p_db[1] + p_hb[1])):
        imgs = [u for u in own if u]
        if not imgs:
            imgs = all_images
        D[k].images = list(dict.fromkeys(imgs))
    missing = [u for u in all_images if u not in D["E"].images and u not in D["B"].images]
    D["E"].images += missing

    for k in "EB":
        D[k].drop_repeated_paragraphs(D[k].bullets)

    return {"H": H, "D": D, "own_desc_bullets": own_desc_bullets, "notes": notes,
            "images": all_images}


def _key(items):
    return "|".join(re.sub(r"\W+", " ", i.lower()).strip() for i in items)


def clean_rows(rows):
    """rows: [(name, he, hb, de, db)] -> ([(he, hb, de, db)], stats Counter, per-row notes)."""
    results = [clean_product(*r) for r in rows]

    # Enrichment: generic highlights reused by 2+ products -> the product's own description bullets.
    stats = Counter()
    for k in "EB":
        h_count = Counter(_key(r["H"][k]) for r in results if r["H"][k])
        d_count = Counter(_key(r["own_desc_bullets"][k]) for r in results if r["own_desc_bullets"][k])
        for r in results:
            own = r["own_desc_bullets"][k]
            if (r["H"][k] and h_count[_key(r["H"][k])] >= 2 and len(own) >= 4
                    and d_count[_key(own)] < 2 and _key(own) != _key(r["H"][k])):
                r["H"][k] = list(own)
                r["notes"].add("highlights_enriched")
                r["D"][k].drop_repeated_paragraphs(r["D"][k].bullets)

    out, notes = [], []
    for r in results:
        out.append((ul(r["H"]["E"]), ul(r["H"]["B"]), r["D"]["E"].render(), r["D"]["B"].render()))
        notes.append(r["notes"])
        stats.update(r["notes"])
    return out, stats, notes


# ---------------------------------------------------------------------------
# Self-verification (the prompt's 6 checks)
# ---------------------------------------------------------------------------
_IMG_SRC = re.compile(r'<img[^>]*\ssrc\s*=\s*(["\'])(.*?)\1', re.IGNORECASE)


def check_row(original_cells, cleaned_cells):
    """Returns a list of problem names (empty list = row passes)."""
    problems = []
    joined = "".join(cleaned_cells)
    if any(real_len(re.sub(r"<[^>]+>", " ", c)) < 3 for c in cleaned_cells):
        problems.append("empty cell")
    if re.search(r"style=|class=|<div|<span|<a\s|<font", joined, re.IGNORECASE):
        problems.append("unwanted markup")
    if re.search(r"<(li|p|ul|h2|h3|td)>\s*</\1>", joined):
        problems.append("empty tag")
    for li in re.findall(r"<li>(.*?)</li>", joined):
        if is_junk_bullet(html.unescape(li)):
            problems.append("junk bullet")
            break
    before = {html.unescape(m[1]).strip() for c in original_cells for m in _IMG_SRC.findall(str(c or ""))}
    after = {html.unescape(m[1]).strip() for c in cleaned_cells for m in _IMG_SRC.findall(c)}
    if before - after:
        problems.append("lost image")
    for desc in cleaned_cells[2:]:
        bullets = " ".join(re.findall(r"<li>(.*?)</li>", desc))
        bw = content_words(html.unescape(bullets))
        for p in re.findall(r"<p>(.*?)</p>", desc):
            pw = content_words(html.unescape(p))
            if pw and len(pw & bw) / len(pw) >= REPEAT_SHARE:
                problems.append("duplicated paragraph")
                break
    return problems
