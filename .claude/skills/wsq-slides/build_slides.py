#!/usr/bin/env python3
"""build_slides.py — WSQ course slide deck for TGS-2024048316 (CompTIA Linux+ XK0-006), v3.

STANDARD HOUSE FORMAT (modeled on the v10 "Mastering the Art of Communication" master
trainer deck): 13.333x7.5in, teal/navy palette, kicker + bold title, rounded accent-bar
cards, dark navy section dividers, 5-step journey cards, right-side imagery.

The deck FOLLOWS THE EXAM DOMAINS objective-by-objective (1.1 → 5.5). The legacy v3
Learner Guide deck's 22 teaching chapters are CONDENSED into standard-format content
slides (key screenshots preserved), slotted under the objective each chapter teaches,
with the 30 hands-on labs given full step-by-step slide coverage. Target 500-700 slides.

Writes courseware/slide_map.json (lab/domain/objective -> deck page) for the Lesson Plan.
"""
import io
import os
import re
import sys
import json

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, os.pardir, "wsq-learner-guide"))
from course_content import (COURSE, DOMAINS, LEARNING_OUTCOMES, DOMAIN_CONCEPTS,
                            DOMAIN_OUTCOMES, domain_labs, COURSEWARE, ASSETS,
                            REFERENCE_DECK, REFERENCE_CHAPTERS, REFERENCE_DROP,
                            OBJ_INFO, CHAPTER_OBJECTIVE, LABS_BY_NUM)

# ---------------- palette (sampled from the v10 standard deck) ----------------
TEAL   = RGBColor(0x00, 0x8C, 0x95)
BLUE   = RGBColor(0x1F, 0x6F, 0xEB)
NAVY   = RGBColor(0x13, 0x1B, 0x2A)   # titles / body ink
DEEP   = RGBColor(0x0E, 0x2A, 0x47)   # dark section background / callout card
GREY   = RGBColor(0x5C, 0x66, 0x77)
CARD   = RGBColor(0xF4, 0xF7, 0xFB)   # card fill
CLINE  = RGBColor(0xD9, 0xE2, 0xEC)   # card outline
CYAN   = RGBColor(0x6D, 0xDA, 0xD7)   # accent on dark
LTXT   = RGBColor(0xD7, 0xE0, 0xEA)   # light text on dark
PURPLE = RGBColor(0x6F, 0x42, 0xC1)
ORANGE = RGBColor(0xE6, 0x83, 0x00)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
CODEBG = RGBColor(0x0B, 0x12, 0x20)
CODEFG = RGBColor(0x9C, 0xDC, 0xFE)
ACCENTS = [BLUE, TEAL, PURPLE, ORANGE]

MAXIM = "Type it, break it, fix it — the exam rewards hands that have done the work."

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

SLIDE_MAP = {"labs": {}, "domains": {}, "objectives": {}, "total": 0}
PAGE = {"n": 0}


# ---------------- primitives ----------------
def slide():
    return prs.slides.add_slide(BLANK)

def rect(s, x, y, w, h, color, line=None, round_=False):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE
    sp = s.shapes.add_shape(shp, x, y, w, h)
    if round_:
        try: sp.adjustments[0] = 0.12
        except Exception: pass
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(0.75)
    sp.shadow.inherit = False
    return sp

def oval(s, x, y, w, h, color):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=4, wrap=True):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    for i, line in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space)
        for spec in line:
            t, sz, col, bold = spec[:4]
            italic = spec[4] if len(spec) > 4 else False
            font = spec[5] if len(spec) > 5 else "Arial"
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.bold = bold; r.font.italic = italic
            r.font.color.rgb = col; r.font.name = font
    return tb

def footer(s, dark=False):
    PAGE["n"] += 1
    col = LTXT if dark else GREY
    txt(s, Inches(0.55), Inches(7.10), Inches(10.8), Inches(0.22),
        [[(f"{COURSE['short']} | {COURSE['code']} | © Tertiary Infotech Academy Pte Ltd", 7.8, col, False)]])
    txt(s, Inches(12.05), Inches(7.08), Inches(0.72), Inches(0.25),
        [[(str(PAGE["n"]), 8, col, False)]], align=PP_ALIGN.RIGHT)
    return PAGE["n"]

def head(s, kicker, title, kcolor=TEAL, tsize=28):
    rect(s, 0, 0, SW, SH, WHITE)
    txt(s, Inches(0.65), Inches(0.34), Inches(11.9), Inches(0.32), [[(kicker.upper(), 9, kcolor, True)]])
    txt(s, Inches(0.65), Inches(0.70), Inches(12.0), Inches(0.90), [[(title, tsize, NAVY, True)]])
    return s

def _logo(name):
    p = os.path.join(ASSETS, name)
    return p if os.path.exists(p) else None

def _fit(text, n):
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= n: return text
    cut = text[:n]
    if " " in cut: cut = cut.rsplit(" ", 1)[0]
    return cut + " …"


# ---------------- components ----------------
def cover():
    s = slide(); rect(s, 0, 0, SW, SH, WHITE)
    rect(s, 0, 0, SW, Inches(0.14), TEAL)
    # right navy panel with course badge
    rect(s, Inches(7.65), Inches(0.14), Inches(5.69), Inches(7.36), DEEP)
    badge = _logo("comptia-linux-logo.png")
    if badge:
        card = rect(s, Inches(8.72), Inches(2.20), Inches(3.55), Inches(2.55), WHITE, round_=True)
        s.shapes.add_picture(badge, Inches(9.22), Inches(2.45), height=Inches(2.05))
    txt(s, Inches(7.65), Inches(5.15), Inches(5.69), Inches(0.5),
        [[("XK0-006  ·  V8  ·  5 DOMAINS  ·  30 HANDS-ON LABS", 11, CYAN, True)]], align=PP_ALIGN.CENTER)
    txt(s, Inches(7.65), Inches(5.62), Inches(5.69), Inches(0.5),
        [[("Powered by Killercoda — every lab free in the browser", 10.5, LTXT, False)]], align=PP_ALIGN.CENTER)
    # left column
    txt(s, Inches(0.72), Inches(0.58), Inches(6.25), Inches(0.32),
        [[("WSQ | COMPTIA LINUX+ CERTIFICATION TRAINING", 10, TEAL, True)]])
    txt(s, Inches(0.72), Inches(1.15), Inches(6.55), Inches(2.25),
        [[("CompTIA Certified Linux+ Training (XK0-006)", 31, DEEP, True)]])
    txt(s, Inches(0.76), Inches(3.45), Inches(6.15), Inches(0.85),
        [[("Administer, secure, automate and troubleshoot Linux — and walk into the exam having already done the work.", 16, GREY, False, True)]])
    txt(s, Inches(0.76), Inches(4.66), Inches(6.30), Inches(0.95),
        [[(f"Course Code: {COURSE['code']}", 12, NAVY, False)],
         [(f"Trainer: {COURSE['trainer']}", 12, NAVY, False)],
         [(f"Version {COURSE['version']} | {COURSE['version_date']}", 12, NAVY, False)]], space=3)
    txt(s, Inches(0.76), Inches(6.23), Inches(6.0), Inches(0.60),
        [[(f"{COURSE['org']} | UEN {COURSE['uen']}", 10, DEEP, True)]])
    footer(s)

def section(kicker, title, sub, maxim=MAXIM, record=None, obj_record=None):
    s = slide(); rect(s, 0, 0, SW, SH, DEEP)
    txt(s, Inches(0.75), Inches(0.72), Inches(6.4), Inches(0.40), [[(kicker.upper(), 11, CYAN, True)]])
    txt(s, Inches(0.72), Inches(1.45), Inches(11.4), Inches(2.15), [[(title, 36, WHITE, True)]])
    txt(s, Inches(0.75), Inches(4.02), Inches(11.4), Inches(1.05), [[(sub, 17, LTXT, False)]])
    txt(s, Inches(0.75), Inches(5.66), Inches(10.4), Inches(0.70), [[(maxim, 15, CYAN, True, True)]])
    p = footer(s, dark=True)
    if record is not None: SLIDE_MAP["domains"][str(record)] = p
    if obj_record is not None: SLIDE_MAP["objectives"][obj_record] = p
    return s

def callout(s, text, y=Inches(1.68), w=Inches(6.28), x=Inches(0.68), h=Inches(1.18)):
    rect(s, x, y, w, h, DEEP, round_=True)
    txt(s, x + Inches(0.30), y + Inches(0.16), w - Inches(0.6), h - Inches(0.32),
        [[(text, 15.5, WHITE, True, True)]], anchor=MSO_ANCHOR.MIDDLE)

def card_rows(s, items, x, y0, w, area_h, size=13):
    """items: list of (text, bold, code). Rounded cards with a coloured accent bar."""
    n = len(items)
    if n == 0: return
    gap = Inches(0.12)
    h = int((area_h - int(gap) * (n - 1)) / n)
    h = min(h, Inches(0.92))
    for i, (text, bold, code) in enumerate(items):
        y = int(y0 + (h + int(gap)) * i)
        rect(s, x, y, w, h, CARD, line=CLINE, round_=True)
        rect(s, x, y, Inches(0.08), h, ACCENTS[i % len(ACCENTS)])
        font = "Consolas" if code else "Arial"
        sz = size - (1.5 if code else 0)
        txt(s, x + Inches(0.30), y, w - Inches(0.55), h,
            [[(text, sz, NAVY, bold, False, font)]], anchor=MSO_ANCHOR.MIDDLE, space=0)

def content_slide(kicker, title, items, image=None, quote=None, record=None):
    """Standard content slide: cards (optionally next to an image), optional navy callout."""
    s = head(slide(), kicker, title)
    has_img = image is not None
    x = Inches(0.68); w = Inches(6.28) if has_img else Inches(12.0)
    y0 = Inches(1.68)
    if quote:
        callout(s, _fit(quote, 150), y=y0, w=w, x=x)
        y0 = Inches(3.05)
    area_h = int(Inches(6.95) - int(y0))
    card_rows(s, items, x, y0, w, area_h)
    if has_img:
        blob, px_w, px_h = image
        box_w, box_h = 5.75, 5.30
        scale = min(box_w / px_w, box_h / px_h)
        iw, ih = px_w * scale, px_h * scale
        ix = 7.35 + (box_w - iw) / 2
        iy = 1.62 + (box_h - ih) / 2
        pic = s.shapes.add_picture(io.BytesIO(blob), Inches(ix), Inches(iy), width=Inches(iw))
        pic.line.color.rgb = CLINE; pic.line.width = Pt(1)
    p = footer(s)
    if record is not None: SLIDE_MAP["labs"][str(record)] = p
    return s

def tile_slide(kicker, title, tiles):
    """2-column grid of mini-cards: (heading, description)."""
    s = head(slide(), kicker, title)
    n = len(tiles); cols = 2 if n > 3 else 1
    rows = -(-n // cols)
    X0 = Inches(0.68); Y0 = Inches(1.75)
    gx = Inches(0.28); gy = Inches(0.20)
    cw = int((Inches(12.0) - int(gx) * (cols - 1)) / cols)
    ch = int((Inches(5.15) - int(gy) * (rows - 1)) / rows)
    for i, (hd, ds) in enumerate(tiles):
        r_, c_ = divmod(i, cols)
        x = int(X0 + (cw + int(gx)) * c_); y = int(Y0 + (ch + int(gy)) * r_)
        rect(s, x, y, cw, ch, CARD, line=CLINE, round_=True)
        rect(s, x, y, Inches(0.08), ch, ACCENTS[i % len(ACCENTS)])
        txt(s, x + Inches(0.28), y + Inches(0.10), cw - Inches(0.5), ch - Inches(0.16),
            [[(hd, 13.5, NAVY, True)], [(ds, 11, GREY, False)]], space=2)
    footer(s)
    return s

def journey(kicker, title, steps, sub=None):
    """Up to 5 rounded step-cards with numbered circles and connector lines."""
    s = head(slide(), kicker, title)
    steps = steps[:5]; n = len(steps)
    y = Inches(2.45); ch = Inches(2.25); cw = Inches(2.19); gap = Inches(0.26)
    tot = int(cw) * n + int(gap) * (n - 1)
    X0 = int((int(SW) - tot) / 2)
    for i, st in enumerate(steps):
        x = int(X0 + (int(cw) + int(gap)) * i)
        rect(s, x, y, cw, ch, CARD, line=CLINE, round_=True)
        oval(s, x + Inches(0.18), y + Inches(0.22), Inches(0.45), Inches(0.45), ACCENTS[i % len(ACCENTS)])
        txt(s, x + Inches(0.18), y + Inches(0.24), Inches(0.45), Inches(0.34),
            [[(str(i + 1), 11, WHITE, True)]], align=PP_ALIGN.CENTER)
        txt(s, x + Inches(0.18), y + Inches(0.88), cw - Inches(0.36), Inches(1.20),
            [[(st, 13, NAVY, True)]], align=PP_ALIGN.CENTER)
        if i < n - 1:
            rect(s, x + int(cw), int(y) + int(ch) // 2, gap, Inches(0.02), CLINE)
    if sub:
        txt(s, Inches(0.68), Inches(5.30), Inches(12.0), Inches(0.6), [[(sub, 13, GREY, False)]], align=PP_ALIGN.CENTER)
    footer(s)
    return s

def quote_dark(big, small, kicker, tag):
    s = slide(); rect(s, 0, 0, SW, SH, RGBColor(0x13, 0x18, 0x27))
    txt(s, Inches(0.78), Inches(0.68), Inches(8.0), Inches(0.35), [[(kicker.upper(), 10, PURPLE, True)]])
    rect(s, Inches(9.55), Inches(0.75), Inches(2.65), Inches(2.45), PURPLE, round_=True)
    txt(s, Inches(10.55), Inches(1.05), Inches(1.6), Inches(1.2), [[("“", 60, WHITE, True)]])
    txt(s, Inches(0.85), Inches(1.85), Inches(8.2), Inches(2.4), [[(big, 30, WHITE, True)]])
    rect(s, Inches(0.85), Inches(4.55), Inches(8.0), Inches(0.09), PURPLE)
    txt(s, Inches(0.87), Inches(4.95), Inches(11.5), Inches(0.8), [[(small, 14, LTXT, False)]])
    txt(s, Inches(0.87), Inches(6.25), Inches(8.0), Inches(0.35), [[(tag.upper(), 10, PURPLE, True)]])
    footer(s, dark=True)
    return s

def trainer_slide(kicker, name, role, rows, initials, accent=TEAL):
    s = head(slide(), kicker, "Your Facilitator")
    lx = Inches(0.68); lw = Inches(3.60)
    rect(s, lx, Inches(1.75), lw, Inches(5.10), CARD, line=CLINE, round_=True)
    rect(s, lx, Inches(1.75), lw, Inches(0.10), accent)
    bd = Inches(1.55); ax = int(lx + (int(lw) - int(bd)) / 2)
    oval(s, ax, Inches(2.20), bd, bd, accent)
    txt(s, ax, Inches(2.60), bd, Inches(0.8), [[(initials, 38, WHITE, True)]], align=PP_ALIGN.CENTER)
    txt(s, lx + Inches(0.15), Inches(4.05), lw - Inches(0.3), Inches(0.55), [[(name, 19, NAVY, True)]], align=PP_ALIGN.CENTER)
    txt(s, lx + Inches(0.15), Inches(4.65), lw - Inches(0.3), Inches(1.2), [[(role, 12, GREY, False)]], align=PP_ALIGN.CENTER)
    rx = Inches(4.70); rw = Inches(7.95)
    items = [(f"{lbl.upper()} — {val}" if val else f"{lbl.upper()} — ____________________", bool(val), False)
             for lbl, val in rows]
    card_rows(s, [(f"{lbl}", False, False) for lbl, _ in []] or items, rx, Inches(1.75), rw, int(Inches(5.10)), size=12)
    footer(s)
    return s

def break_slide(kind, dur):
    s = slide(); rect(s, 0, 0, SW, SH, WHITE)
    rect(s, 0, 0, SW, Inches(0.14), TEAL); rect(s, 0, Inches(7.36), SW, Inches(0.14), TEAL)
    rect(s, Inches(5.4), Inches(2.45), Inches(2.53), Inches(0.09), TEAL)
    txt(s, 0, Inches(2.85), SW, Inches(1.1), [[(kind, 44, NAVY, True)]], align=PP_ALIGN.CENTER)
    txt(s, 0, Inches(4.05), SW, Inches(0.7), [[(dur, 20, TEAL, True)]], align=PP_ALIGN.CENTER)
    PAGE["n"] += 1


# ---------------- legacy content extraction (condensed) ----------------
_MONO = re.compile(r"mono|courier|consolas", re.I)

def _parse_slide(s):
    title = ""; bullets = []; best_img = None; best_area = 0
    for sh in s.shapes:
        if sh.shape_type == 13:
            area = int(sh.width) * int(sh.height)
            if area > best_area:
                try:
                    best_img = sh.image.blob; best_area = area
                except Exception:
                    pass
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if not t or t == "‹#›" or t.startswith("This material belongs"):
                continue
            paras = sh.text_frame.paragraphs
            start = 0
            if not title:
                title = paras[0].text.strip() if paras else t.split("\n")[0]
                start = 1
            for p in paras[start:]:
                line = p.text.strip()
                if len(line) <= 3: continue
                # drop orphan fragments (bare labels like "State Files:")
                if len(line) < 20 and line.endswith(":"): continue
                code = any(_MONO.search(r.font.name or "") for r in p.runs if r.text.strip())
                bullets.append((line, code))
    return title, bullets, best_img, best_area

def load_condensed():
    """Parse the legacy deck into per-objective, ordered, deduped topic groups."""
    legacy = Presentation(REFERENCE_DECK)
    lslides = list(legacy.slides)
    obj_groups = {code: [] for code in OBJ_INFO}
    for ch, a, b, dom, name in REFERENCE_CHAPTERS:
        obj = CHAPTER_OBJECTIVE.get(ch)
        if obj is None: continue
        groups = {}; order = []
        for i in range(a, b + 1):
            if i in REFERENCE_DROP: continue
            t, bl, img, area = _parse_slide(lslides[i - 1])
            key = re.sub(r"\s+", " ", t).lower().strip()
            if not key: continue
            # skip the legacy chapter-header slides — structural, not content
            if re.match(r"^(chapter|part)\s+\d+", key): continue
            if key not in groups:
                groups[key] = dict(title=t.strip(), bullets=[], img=None, img_area=0, chapter=name)
                order.append(key)
            g = groups[key]
            seen = {x[0].lower() for x in g["bullets"]}
            for text, code in bl:
                if text.lower() not in seen:
                    g["bullets"].append((text, code)); seen.add(text.lower())
            # keep only screenshots of a meaningful size (> ~2 sq in on the old canvas)
            if img is not None and area > 914400 * 914400 * 2 and area > g["img_area"]:
                g["img"], g["img_area"] = img, area
        for k in order:
            obj_groups[obj].append(groups[k])
    return obj_groups

def emit_condensed(obj, groups, budget):
    """Emit at most `budget` condensed slides for one objective's topic groups."""
    if not groups or budget <= 0: return 0
    cands = [(i, g) for i, g in enumerate(groups) if g["img"] is not None or len(g["bullets"]) >= 5]
    small = [(i, g) for i, g in enumerate(groups) if (i, g) not in cands]
    tile_slides = min(2, -(-len(small) // 6)) if small else 0
    max_own = max(1, budget - tile_slides)
    if len(cands) > max_own:
        ranked = sorted(cands, key=lambda ig: (ig[1]["img"] is not None, len(ig[1]["bullets"])), reverse=True)
        keep = {id(g) for _, g in ranked[:max_own]}
        demoted = [(i, g) for i, g in cands if id(g) not in keep]
        cands = [(i, g) for i, g in cands if id(g) in keep]
        small = sorted(small + demoted, key=lambda ig: ig[0])
        tile_slides = min(2, -(-len(small) // 6)) if small else 0
    n_emitted = 0
    oname, _ = OBJ_INFO[obj]
    kick = f"OBJECTIVE {obj} | {oname}"
    for _, g in cands:
        items = [(_fit(t, 155), i == 0, c) for i, (t, c) in enumerate(g["bullets"][:6])]
        if not items:
            items = [(g["chapter"], False, False)]
        img = None
        if g["img"] is not None:
            try:
                from pptx.parts.image import Image as _PImg
                pim = _PImg.from_blob(g["img"])
                img = (g["img"], pim.size[0], pim.size[1])
            except Exception:
                img = None
        content_slide(kick, _fit(g["title"], 70), items, image=img)
        n_emitted += 1
    # remaining topics -> compact 6-tile digests (condensation)
    emitted_tiles = 0
    for t_i in range(tile_slides):
        chunk = small[t_i * 6:(t_i + 1) * 6]
        if not chunk: break
        tiles = [(_fit(g["title"], 55), _fit(" ".join(x[0] for x in g["bullets"][:2]) or g["chapter"], 130))
                 for _, g in chunk]
        tile_slide(kick, f"{oname} — More Key Topics", tiles)
        n_emitted += 1; emitted_tiles += 1
    return n_emitted


# ---------------- labs ----------------
def first_cmd(code):
    for ln in code.split("\n"):
        t = ln.strip()
        if t and not t.startswith("#") and not t.startswith("cat <<") and "EOF" not in t:
            return t[:80]
    return code.split("\n")[0][:80]

def lab_slides(lab):
    n = lab["num"]; obj = lab["objective"]
    kick = f"LAB {n} | OBJECTIVE {obj} | HANDS-ON"
    # overview
    s = head(slide(), kick, _fit(f"Lab {n} — {lab['title']}", 75))
    callout(s, _fit(lab["goal"], 165), w=Inches(12.0))
    items = [("You'll build — " + _fit(lab["build"], 120), True, False),
             ("Key commands — " + _fit(", ".join(lab["concepts"][:6]), 120), False, False),
             ("Workbench — Killercoda Ubuntu Playground, free in the browser", False, False),
             ("Full walkthrough — labs/lab-%02d-*/README.md on the course GitHub repo" % n, False, False)]
    card_rows(s, items, Inches(0.68), Inches(3.10), Inches(12.0), int(Inches(3.80)))
    p = footer(s)
    SLIDE_MAP["labs"][str(n)] = p
    # steps, 4 per slide, verify appended as the final card
    steps = lab["steps"]
    cards = [(f"Step {i} — " + _fit(st["title"], 90), first_cmd(st["code"])) for i, st in enumerate(steps, 1)]
    cards.append(("✅ Verify — " + _fit(lab["test"], 110), ""))
    per = 4
    pages = [cards[i:i + per] for i in range(0, len(cards), per)]
    for pi, chunk in enumerate(pages, 1):
        s = head(slide(), f"LAB {n} | STEP-BY-STEP ({pi}/{len(pages)})", _fit(lab["title"], 70), kcolor=BLUE)
        y0 = Inches(1.75); gap = Inches(0.16)
        h = int((Inches(5.05) - int(gap) * (len(chunk) - 1)) / len(chunk))
        h = min(h, Inches(1.24))
        for i, (t, cmd) in enumerate(chunk):
            y = int(y0 + (h + int(gap)) * i)
            rect(s, Inches(0.68), y, Inches(12.0), h, CARD, line=CLINE, round_=True)
            rect(s, Inches(0.68), y, Inches(0.08), h, ACCENTS[i % len(ACCENTS)])
            if cmd:
                txt(s, Inches(0.98), y + Inches(0.08), Inches(11.5), Inches(0.42),
                    [[(t, 12.5, NAVY, True)]], space=0)
                cy = y + int(h) - int(Inches(0.52))
                rect(s, Inches(0.98), cy, Inches(11.35), Inches(0.40), CODEBG)
                txt(s, Inches(1.18), cy + Inches(0.03), Inches(11.0), Inches(0.34),
                    [[("$ " + cmd, 10.5, CODEFG, False, False, "Consolas")]], space=0, wrap=False)
            else:
                txt(s, Inches(0.98), y, Inches(11.5), h,
                    [[(t, 13, NAVY, True)]], anchor=MSO_ANCHOR.MIDDLE, space=0)
        footer(s)


# ============================================================ BUILD
COND = load_condensed()
TOTAL_GROUPS = sum(len(v) for v in COND.values())
CONTENT_BUDGET = 360

cover()

# ---------------- SECTION 00 — course administration ----------------
section("SECTION 00", "Course Administration",
        "Welcome | attendance | facilitator | assessment | resources")
content_slide("MANDATORY DIGITAL ATTENDANCE", "Digital Attendance (TRAQOM · SSG)",
    [("It is mandatory to take the AM, PM and Assessment digital attendance for WSQ-funded courses.", True, False),
     ("The trainer/administrator displays the digital attendance QR code from the SSG portal.", False, False),
     ("Scan the QR code with your mobile phone camera and submit your attendance.", False, False),
     ("A minimum of 75% attendance is required to be eligible for assessment and funding.", False, False)])
trainer_slide("YOUR FACILITATOR | GENERAL", "Your Trainer",
    "General Trainer template —\nto be completed by the trainer",
    [("Name", ""), ("Title / Designation", ""), ("Qualifications", ""),
     ("Areas of expertise", ""), ("Training & industry experience", ""), ("Contact", "")],
    initials="?", accent=GREY)
trainer_slide("YOUR FACILITATOR", COURSE["trainer"],
    "Principal Trainer\nTertiary Infotech Academy Pte. Ltd.",
    [("Role", "Principal Trainer, Tertiary Infotech Academy Pte. Ltd."),
     ("Focus", "Linux, cloud, DevOps & cybersecurity — hands-on systems administration"),
     ("Delivers", "WSQ courses on Linux, cloud administration, DevOps and software engineering"),
     ("Founder", "Founder and lead instructor, Tertiary Infotech / Tertiary Courses")],
    initials="AA", accent=TEAL)
content_slide("LEARNER INTRODUCTION", "Let's Know Each Other",
    [("Your name and organisation / role.", True, False),
     ("Your current experience with Linux — beginner to advanced.", False, False),
     ("Which of the five exam domains matters most to your day job.", False, False),
     ("Your CompTIA Linux+ exam timeline.", False, False)])
tile_slide("HOW WE WILL WORK", "Ground Rules",
    [("Phones on silent", "Keep the room focused — check messages at breaks."),
     ("Participate actively", "No question is too small; the discussion is the learning."),
     ("Do every lab yourself", "Type the commands on Killercoda — muscle memory wins exams."),
     ("Mutual respect", "Agree to disagree; critique ideas, not people."),
     ("Be punctual", "Return from breaks on time so nobody re-teaches."),
     ("75% attendance", "Required for assessment eligibility and WSQ funding.")])
journey("COURSE OUTLINE", "Two-Day Learning Journey",
    ["System Management", "Services & Users", "Security", "Automation & Scripting", "Troubleshooting"],
    sub="Day 1: Domains 1–2 (Labs 1–13)  ·  Day 2: Domains 3–5 (Labs 14–30) + Final Assessment  ·  9:30am–6:30pm")
tile_slide("WHAT YOU WILL BE ABLE TO DO", "Learning Outcomes",
    [(f"LO{i} — {DOMAIN_OUTCOMES[i][0]}", DOMAIN_OUTCOMES[i][1]) for i in range(1, 6)])
tile_slide("EXAM BLUEPRINT | XK0-006 V8", "CompTIA Linux+ — Exam Domains",
    [(f"Domain {d['num']} — {d['title']} ({d['weight']}%)",
      f"Objectives {d['objs'][0][0]}–{d['objs'][-1][0]}  ·  Labs {d['labs'][0]}–{d['labs'][-1]}")
     for d in DOMAINS])
content_slide("EXAM-ALIGNED STRUCTURE", "How This Deck Is Organised",
    [("The course follows the five XK0-006 exam domains in blueprint order.", True, False),
     ("Inside each domain, content runs objective by objective (1.1 → 5.5).", False, False),
     ("Each objective: the concepts, condensed — then its hands-on lab, step by step.", False, False),
     ("Every lab maps to one exam objective, shown on its overview slide.", False, False),
     ("The deck closes with exam registration, the practice exam and the WSQ assessment.", False, False)])
tile_slide("ASSESSMENT READINESS", "Assessment — What to Expect",
    [("Written Assessment (WA)", "Short-Answer Questions (SAQ) · 1 hour · open book · aligned to the slides."),
     ("Practical Performance (PP)", "Hands-on Linux tasks · 1 hour · open book · aligned to the labs."),
     ("Open book", "Slides, Learner Guide and approved materials only."),
     ("Eligibility", "Minimum 75% attendance · an appeal process is available.")])
tile_slide("BEFORE THE ASSESSMENT", "Briefing for Assessment",
    [("Clear your table", "Phones and materials under the table or on the floor."),
     ("No photos / recording", "Assessment scripts must not be photographed or recorded."),
     ("No discussion", "Work alone during the assessment."),
     ("Blue / black pen", "For hard-copy assessments; no liquid paper or correction tape."),
     ("Time's up = pens down", "Scripts are collected when time is up."),
     ("Submit on the LMS", "Upload your completed answers to lms-tms.tertiaryinfotech.com.")])
journey("ASSESSMENT PROCESS", "On Assessment Day",
    ["TRAQOM digital attendance", "Assessment attendance", "Sit WA (SAQ), then PP", "Submit on the LMS", "Sign the Summary Record"])
# download course material — visual with the LMS screenshot
_shot = _logo("lms-tms-portal.png")
_dl_items = [("1 — Go to the portal:  lms-tms.tertiaryinfotech.com", True, False),
             ("2 — Log in with your registered email (Send OTP or password)", False, False),
             ("3 — Open your course: CompTIA Certified Linux+ Training", False, False),
             ("4 — Download the slides, Learner Guide and Lesson Plan (open-book allowed)", False, False)]
if _shot:
    from pptx.parts.image import Image as _PImg2
    with open(_shot, "rb") as fh: _blob = fh.read()
    _pim = _PImg2.from_blob(_blob)
    content_slide("LEARNER RESOURCES | LMS / TMS", "Download Course Material", _dl_items,
                  image=(_blob, _pim.size[0], _pim.size[1]))
else:
    content_slide("LEARNER RESOURCES | LMS / TMS", "Download Course Material", _dl_items)
content_slide("ACCESS THE LABS | GITHUB", "Access the Hands-On Labs",
    [("All 30 labs are on the course GitHub repository — one folder per lab:", True, False),
     (COURSE["repo"], False, True),
     ("Option A — git clone, then open labs/lab-XX-*/README.md", False, False),
     ("Option B — Code ▸ Download ZIP, unzip, open the labs/ folder", False, False),
     ("Every lab runs free in the browser on Killercoda: " + COURSE["killercoda"], False, False)])
quote_dark("On Linux, everything is a file.",
           "Files, directories, devices, sockets and processes are all reachable through the filesystem — "
           "master the shell and you master the system.",
           "THE LINUX MINDSET", "COURSE MAXIM")

# ---------------- DOMAINS ----------------
BREAKS = {1: ("Tea Break", "15 minutes"), 2: ("End of Day 1", "See you tomorrow — 9:30 AM"),
          3: ("Lunch Break", "1 hour"), 4: ("Tea Break", "15 minutes")}
for d in DOMAINS:
    dn = d["num"]
    section(f"DOMAIN 0{dn}", d["title"],
            f"{d['weight']}% of the exam | Objectives {d['objs'][0][0]}–{d['objs'][-1][0]} | Labs {d['labs'][0]}–{d['labs'][-1]}",
            record=dn)
    tile_slide(f"DOMAIN 0{dn} | EXAM OBJECTIVES", f"{d['title']} — Objectives",
        [(f"{code}  ·  {OBJ_INFO[code][0]}", OBJ_INFO[code][1]) for code, _ in d["objs"]])
    content_slide(f"DOMAIN 0{dn} | KEY CONCEPTS", f"Key Concepts — {d['title']}",
        [(_fit(c, 165), i == 0, False) for i, c in enumerate(DOMAIN_CONCEPTS[dn])])
    labs_by_obj = {}
    for lab in domain_labs(dn):
        labs_by_obj.setdefault(lab["objective"], []).append(lab)
    for code, formal in d["objs"]:
        oname, osum = OBJ_INFO[code]
        obj_lab_names = ", ".join(f"Lab {l['num']}" for l in labs_by_obj.get(code, [])) or "—"
        s = head(slide(), f"DOMAIN 0{dn} | OBJECTIVE {code}", f"{code} — {oname}", kcolor=BLUE, tsize=26)
        callout(s, osum, w=Inches(12.0))
        card_rows(s, [("Exam objective — " + formal, True, False),
                      ("Hands-on — " + obj_lab_names, False, False)],
                  Inches(0.68), Inches(3.10), Inches(12.0), int(Inches(1.90)))
        txt(s, Inches(0.68), Inches(5.35), Inches(12.0), Inches(1.3),
            [[(MAXIM, 13, TEAL, True, True)]])
        p = footer(s)
        SLIDE_MAP["objectives"].setdefault(code, p)
        groups = COND.get(code, [])
        budget = max(2, round(CONTENT_BUDGET * len(groups) / TOTAL_GROUPS)) if groups else 0
        emit_condensed(code, groups, budget)
        for lab in labs_by_obj.get(code, []):
            lab_slides(lab)
    # labs not tied to a single objective (e.g. the capstone) close out the domain
    _codes = {code for code, _ in d["objs"]}
    for lab in domain_labs(dn):
        if lab["objective"] not in _codes:
            lab_slides(lab)
    content_slide(f"DOMAIN 0{dn} | RECAP", f"Recap — {d['title']}",
        [(f"Lab {l['num']} — " + _fit(l["title"], 90), False, False) for l in domain_labs(dn)][:7])
    if dn in BREAKS:
        break_slide(*BREAKS[dn])

# ---------------- WRAP-UP & CERTIFICATION ----------------
section("SECTION 06", "Wrap-Up & Certification",
        "Summary | exam registration | practice exam | assessment")
tile_slide("WHAT YOU ACHIEVED", "Learning Outcomes — Revisited",
    [(f"LO{i} — {DOMAIN_OUTCOMES[i][0]}", DOMAIN_OUTCOMES[i][1]) for i in range(1, 6)])
content_slide("NEXT STEPS", "Preparing for the Linux+ XK0-006 Exam",
    [("Redo every lab on Killercoda until the commands are automatic.", True, False),
     ("Review each lab's Verify takeaway and the Learner Guide.", False, False),
     ("Know which tool solves which problem — the exam is scenario-based.", False, False),
     ("Up to 90 questions (multiple-choice + performance-based) in 90 minutes; pass 720/900.", False, False),
     ("Book through Pearson VUE from your CompTIA account.", False, False)])
journey("GET CERTIFIED | XK0-006", "CompTIA Linux+ Exam Registration",
    ["Create your CompTIA account", "Buy the XK0-006 voucher", "Schedule at Pearson VUE", "Sit the exam — 90 min", "Claim your digital badge"],
    sub=f"Voucher: {COURSE['exam_voucher']}   ·   Schedule: {COURSE['exam_pearson']}")
tile_slide("REGISTER FOR THE EXAM", "Exam Registration — Links & Details",
    [("Exam voucher", f"CompTIA Store: {COURSE['exam_voucher']}  (or through Tertiary Infotech)"),
     ("Schedule the exam", f"Pearson VUE: {COURSE['exam_pearson']} — test centre or online proctored"),
     ("Format", "XK0-006 · up to 90 questions · 90 minutes · passing score 720 (scale 100–900)"),
     ("Question types", "Multiple-choice + performance-based items (live Linux tasks)"),
     ("On exam day", "Two forms of ID at a test centre · check online-testing system requirements for remote")])
tile_slide("TEST YOURSELF | EXAMS.TERTIARYINFOTECH.COM", "Practice Exam",
    [("Sharpen your readiness", "Attempt the Tertiary Infotech CompTIA Linux+ practice exam under timed conditions."),
     ("Practice exam link", COURSE["practice_exam"]),
     ("Review every explanation", "Wrong answers are the syllabus for your next study session."),
     ("Close the gaps", "Revisit any lab whose domain you miss, then re-take the practice exam.")])
tile_slide("WRAP-UP | FINAL ASSESSMENT", "Assessment",
    [("Two instruments", "Written Assessment (SAQ) — 1 hour · Practical Performance (PP) — 1 hour."),
     ("Open book", "Slides, Learner Guide and approved materials only."),
     ("Digital attendance", "Remember the Assessment digital attendance (TRAQOM · SSG)."),
     ("Submission", f"Submit your completed answers on the LMS: {COURSE['lms']}")])
journey("ON ASSESSMENT DAY", "Assessment Process",
    ["TRAQOM digital attendance", "Assessment attendance", "Sit WA (SAQ), then PP", "Submit on the LMS", "Sign the Summary Record"])
content_slide("TRAQOM | SSG DIGITAL ATTENDANCE", "Digital Attendance (Mandatory)",
    [("Take the AM, PM and Assessment digital attendance — mandatory for WSQ funding.", True, False),
     ("Scan the SSG QR code displayed by the trainer/administrator.", False, False),
     ("Submit your attendance before leaving each session.", False, False),
     ("75% minimum attendance for assessment eligibility and funding.", False, False)])
quote_dark("You are ready to administer Linux.",
           "Thank you — and all the best for the CompTIA Linux+ XK0-006 exam. "
           f"Register for the exam, take the practice test, and keep every lab within arm's reach: {COURSE['repo']}",
           "THANK YOU", "SEE YOU AT THE EXAM")

# ---------------- save ----------------
SLIDE_MAP["total"] = PAGE["n"]
assert len(prs.slides._sldIdLst) == PAGE["n"], f"{len(prs.slides._sldIdLst)} vs {PAGE['n']}"
OUT = os.path.join(COURSEWARE, f"PPT-CompTIA-Linux-Plus-XK0-006-{COURSE['version']}.pptx")
prs.save(OUT)
with open(os.path.join(COURSEWARE, "slide_map.json"), "w") as fh:
    json.dump(SLIDE_MAP, fh, indent=2)
print(f"Saved {OUT}  ({PAGE['n']} slides)")
print("Wrote slide_map.json")
