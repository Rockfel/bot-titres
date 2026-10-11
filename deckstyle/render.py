"""Mise en page et génération du .pptx au design système (Poppins, blanc, #122F4C, fond dégradé)."""
from __future__ import annotations

import copy
import os
import re

from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
TEMPLATE = os.path.join(ASSETS, "template.pptx")

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x12, 0x2F, 0x4C)

# Une seule famille « Poppins » (Regular / Italic / Bold) : c'est le nom que PowerPoint, Keynote
# et Canva reconnaissent partout. Les graisses Light/Medium sont souvent mal reprises à l'import.
FONT_FILES = {"L": "Poppins-Regular.ttf", "LI": "Poppins-Italic.ttf", "M": "Poppins-Bold.ttf", "B": "Poppins-Bold.ttf"}
FONTS = {k: ImageFont.truetype(os.path.join(ASSETS, "fonts", v), 100) for k, v in FONT_FILES.items()}

FACE = {"L": ("Poppins", False, False), "LI": ("Poppins", False, True),
        "M": ("Poppins", True, False), "B": ("Poppins", True, False)}

# styles : taille (pt), graisse, interligne (x taille), espacement des lettres (pt), majuscules
STY = {
    "eyebrow":  dict(size=15, st="B", lh=1.35, cs=1.4, caps=True),
    "ref":      dict(size=16, st="B", lh=1.35, cs=1.0, caps=True),
    "section":  dict(size=18, st="B", lh=1.35, cs=1.2, caps=True),
    "cover":    dict(size=46, st="M", lh=1.2),
    "hero":     dict(size=38, st="L", lh=1.25),
    "stmt":     dict(size=28, st="L", lh=1.35),
    "stmt_l":   dict(size=32, st="L", lh=1.3),
    "intro":    dict(size=20, st="LI", lh=1.4),
    "big":      dict(size=60, st="L", lh=1.15),
    "quote":    dict(size=24, st="L", lh=1.45),
    "quote_l":  dict(size=28, st="L", lh=1.4),
    "quote_xl": dict(size=36, st="L", lh=1.3),
    "label":    dict(size=16, st="LI", lh=1.3),
    "q":        dict(size=28, st="B", lh=1.3),
    "a":        dict(size=28, st="L", lh=1.3),
}

LAYOUT = {"cover": "DECLARATION", "statement": "DECLARATION", "reading": "LECTURE", "plan": "LECTURE",
          "question": "INTERPELLATION", "twocol": "DEUX_COLONNES"}


# ------------------------------------------------------------------ typographie
def nb(t: str) -> str:
    """Espaces insécables de la typographie française."""
    t = re.sub(r"«\s+", "« ", t)
    t = re.sub(r"\s+»", " »", t)
    return re.sub(r" ([;:?!])", " \\1", t)


def tw(text, st, size, cs=0.0):
    return FONTS[st].getlength(text) * size / 100 + cs * len(text)


def count_lines(runs, size, width, cs=0.0):
    lines, cur = 1, 0.0
    for text, st in runs:
        parts = text.split(" ")
        for i, word in enumerate(parts):
            w = tw(word, st, size, cs)
            sp = tw(" ", st, size, cs) if i < len(parts) - 1 else 0
            if cur > 0 and cur + w > width:
                lines += 1
                cur = 0.0
            cur += w + sp
    return lines


def split_emphasis(text: str, emphasis: list[str]) -> list[tuple[str, bool]]:
    """Découpe `text` en segments (texte, en_valeur)."""
    marks = [False] * len(text)
    for e in emphasis or []:
        e = e.strip()
        if not e:
            continue
        start = text.find(e)
        if start < 0:
            # tolère les espaces insécables / différences d'espaces
            pat = re.escape(e).replace(r"\ ", r"[\s ]+")
            m = re.search(pat, text)
            if not m:
                continue
            start, end = m.span()
        else:
            end = start + len(e)
        for i in range(start, end):
            marks[i] = True
    segs = []
    for ch, m in zip(text, marks):
        if segs and segs[-1][1] == m:
            segs[-1] = (segs[-1][0] + ch, m)
        else:
            segs.append((ch, m))
    return segs or [("", False)]


# ------------------------------------------------------------------ choix des styles
def style_for(mode: str, btype: str, slide: dict) -> str | None:
    quotes = sum(len(b["text"]) for b in slide.get("blocks", []) if b["type"] == "quote")
    stmts = sum(len(b["text"]) for b in slide.get("blocks", []) if b["type"] in ("statement", "title"))
    if btype == "eyebrow":
        return "eyebrow"
    if btype == "ref":
        return "ref"
    if btype == "section":
        return "section"
    if btype == "intro":
        return "intro"
    if btype == "big":
        return "big"
    if btype == "label":
        return "label"
    if mode == "cover":
        return {"title": "cover", "statement": "hero"}.get(btype, "hero")
    if mode == "question":
        return "stmt_l" if stmts < 260 else "stmt"
    if mode == "reading":
        if btype == "question":
            return "q"
        if btype == "answer":
            return "a"
        return "quote_l" if quotes < 300 else "quote"
    # statement
    if btype == "title":
        return "hero"
    if btype == "statement":
        return "stmt_l" if stmts < 220 else "stmt"
    if btype == "question":
        return "q"
    if btype == "answer":
        return "a"
    if btype == "quote":
        return "quote_xl" if quotes < 120 else "quote_l"
    return "stmt"


# ------------------------------------------------------------------ grille et espacements
# Grille de 8 pt. Toutes les slides partagent la même marge gauche (88 pt) et la même
# largeur de texte (784 pt), qu'il y ait une carte ou non : les textes s'alignent d'une
# slide à l'autre. Les cartes ont 48 pt de marge intérieure.
SP = {"xs": 8, "s": 16, "m": 24, "l": 32, "xl": 48}
CONTENT_X, CONTENT_W = 88, 784
CARD = (40, 40, 880, 460)                     # x, y, w, h (slide 960 x 540)
CARD_PAD = 48
TWO_COLS = ((40, 80, 430, 380), (490, 80, 430, 380))
TWO_PAD = 40

# cartes dessinées (le gabarit .pptx contient les mêmes) : fond, opacité, contour
CARDS = {
    "reading":  [dict(rect=CARD, fill="122F4C", alpha=0.78)],
    "plan":     [dict(rect=CARD, fill="122F4C", alpha=0.78)],
    "question": [dict(rect=CARD, fill="0E2740", alpha=0.94, line="FFFFFF", line_alpha=0.45)],
    "twocol":   [dict(rect=r, fill="122F4C", alpha=0.78) for r in TWO_COLS],
}


def gap(prev: str | None, cur: str) -> float:
    """Espace (pt) entre deux blocs, pris dans l'échelle 8/16/24/32/48."""
    if prev is None:
        return 0
    if cur == "question":
        return 40
    table = {
        ("dash", None): SP["m"],
        ("eyebrow", "dash"): SP["s"], ("section", "dash"): SP["s"],
        ("eyebrow", None): SP["m"], ("section", None): SP["l"],
        ("ref", None): SP["xs"], ("label", None): 4,
        ("quote", "quote"): SP["s"], ("quote", "ref"): SP["l"], ("quote", "label"): SP["m"],
        ("question", "answer"): SP["xs"],
        ("intro", None): SP["l"],
        ("big", "big"): SP["xs"],
        ("statement", "statement"): SP["s"],
        ("title", "title"): SP["l"],
    }
    return table.get((prev, cur), table.get((prev, None), SP["m"]))


def _snap(v: float) -> float:
    return round(v / 2) * 2


# ------------------------------------------------------------------ mise en page -> opérations
def _text_op(x, y, w, runs, size, lh, cs=0.0, color="FFFFFF", name="Texte"):
    n = count_lines([(t, st) for t, st, *_ in runs], size, w - 6, cs)
    return dict(kind="text", x=x, y=_snap(y), w=w, h=n * lh + 2, runs=runs, size=size, lh=lh, cs=cs, color=color, name=name)


def _flow(slide_spec):
    """Positionne les blocs sur la grille ; réduit l'ensemble si ça ne tient pas."""
    mode = slide_spec["mode"]
    in_card = mode in ("reading", "question")
    x, w = CONTENT_X, CONTENT_W
    if in_card:
        top, bottom = CARD[1] + CARD_PAD, CARD[1] + CARD[3] - CARD_PAD
    else:
        top, bottom = 72, 468
    blocks = [b for b in slide_spec.get("blocks", []) if b.get("type") != "footer"]
    scale = 1.0
    while True:
        items, cur, prev = [], 0.0, None
        for b in blocks:
            t = b["type"]
            cur += gap(prev, t) * scale
            prev = t
            if t == "dash":
                items.append(dict(kind="dash", x=x, y=cur, w=40, h=3)); cur += 3; continue
            sname = style_for(mode, t, slide_spec)
            st = STY[sname]
            size = round(st["size"] * scale, 1)
            text = nb(b.get("text", "").strip())
            emph = b.get("emphasis") or []
            if st.get("caps"):
                text, emph = text.upper(), [e.upper() for e in emph]
            base = st["st"]
            if sname == "q" or (t == "question" and mode == "question"):
                runs = [(text, "B")]
            else:
                runs = [(seg, "B" if m else base) for seg, m in split_emphasis(text, [nb(e) for e in emph])]
            lh = size * st["lh"]
            op = _text_op(x, cur, w, runs, size, lh, st.get("cs", 0), name=t)
            items.append(op)
            cur += op["h"] - 2
        total = cur
        if total <= bottom - top or scale <= 0.62:
            break
        scale -= 0.04
    if mode == "cover":
        start = top + SP["s"]
    elif in_card:
        start = (CARD[1] + CARD[3] / 2) - total / 2        # centré dans la carte
    else:
        start = 258 - total / 2                            # centre optique, un peu au-dessus du milieu
    start = max(top, min(start, bottom - total))
    for it in items:
        it["y"] = _snap(it["y"] + start)
    return items


def layout_slide(spec) -> list[dict]:
    """Liste d'opérations de dessin (texte, trait) pour une slide, communes au .pptx et au .pdf."""
    mode = spec["mode"]
    ops = []
    if mode == "plan":
        rows = spec.get("rows") or []
        x0, y0 = CARD[0] + CARD_PAD, CARD[1] + CARD_PAD
        ops.append(_text_op(x0, y0, CONTENT_W, [("Plan", "B")], 40, 48, name="Titre"))

        def plan_runs(r):
            runs = [(nb(r.get("text", "")), "L")]
            if r.get("ref"):
                runs.append(("  ·  " + nb(r["ref"]), "B", 75))  # références toujours en gras
            return runs
        num_w = 56
        size = 24
        while True:
            lines = [count_lines([(t, st) for t, st, *_ in plan_runs(r)], size, CONTENT_W - num_w - 6) for r in rows]
            total = sum(n * size * 1.35 for n in lines) + 12 * max(0, len(rows) - 1)
            if total <= CARD[1] + CARD[3] - CARD_PAD - (y0 + 48 + SP["m"]) or size <= 16:
                break
            size -= 1
        y = y0 + 48 + SP["m"]
        for i, (r, n) in enumerate(zip(rows, lines), 1):
            lh = size * 1.35
            ops.append(_text_op(x0, y, num_w, [(r.get("num", ""), "B")], size * 0.85, lh, name=f"Num {i}"))
            ops.append(_text_op(x0 + num_w, y, CONTENT_W - num_w, plan_runs(r), size, lh, name=f"Partie {i}"))
            y += n * lh + 12
        return ops

    if mode == "twocol":
        cols = (spec.get("cols") or [])[:2]
        heads = [nb(c.get("head", "")) for c in cols]
        bodies = [nb(c.get("body", "")) for c in cols]
        cw = TWO_COLS[0][2] - 2 * TWO_PAD
        size = 20
        while size > 13 and max(count_lines([(b, "L")], size, cw - 6) for b in bodies) * size * 1.5 > TWO_COLS[0][3] - 2 * TWO_PAD - 40 - SP["s"]:
            size -= 0.5
        for k, (h, b) in enumerate(zip(heads, bodies)):
            cx, cy = TWO_COLS[k][0] + TWO_PAD, TWO_COLS[k][1] + TWO_PAD
            ops.append(_text_op(cx, cy, cw, [(h, "B")], 28, 36, name=f"Titre colonne {k + 1}"))
            ops.append(_text_op(cx, cy + 36 + SP["s"], cw, [(b, "L")], size, size * 1.5, name=f"Texte colonne {k + 1}"))
        return ops

    ops = _flow(spec)
    footer = [b for b in spec.get("blocks", []) if b.get("type") == "footer"]
    if footer:
        ops.append(_text_op(CONTENT_X, 468 - 22, CONTENT_W, [(nb(footer[0]["text"]), "L")], 16, 22, name="Pied de page"))
    return ops


# ------------------------------------------------------------------ écriture .pptx
def _textbox(slide, x, y, w, h, runs, size, lh, cs, color, name, alpha=None):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.line_spacing = Pt(lh)
    for text, st, *rest in runs:
        r = p.add_run()
        r.text = text
        face, bold, italic = FACE[st]
        f = r.font
        f.name, f.size, f.bold, f.italic = face, Pt(size), bold, italic
        f.color.rgb = color
        rpr = r._r.get_or_add_rPr()
        if cs:
            rpr.set("spc", str(int(cs * 100)))
        run_alpha = rest[0] if rest else alpha
        if run_alpha is not None:
            clr = rpr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(run_alpha * 1000))}))
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set("typeface", face)
    return tb


def _dash(slide, x, y, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.name = "Trait"
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False


def _normalize(slides):
    for spec in slides:
        if spec.get("mode") not in LAYOUT:
            spec["mode"] = "statement"
    return slides


def build(slides: list[dict], out_path: str) -> str:
    prs = Presentation(TEMPLATE)
    layouts = {l.name: l for l in prs.slide_layouts}
    for spec in _normalize(slides):
        slide = prs.slides.add_slide(layouts[LAYOUT[spec["mode"]]])
        for op in layout_slide(spec):
            if op["kind"] == "dash":
                _dash(slide, op["x"], op["y"], op["w"], op["h"], WHITE)
            else:
                _textbox(slide, op["x"], op["y"], op["w"], op["h"], op["runs"], op["size"], op["lh"], op["cs"],
                         RGBColor.from_string(op["color"]), op["name"])
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------------ écriture .pdf (Poppins incorporée)
def build_pdf(slides: list[dict], out_path: str) -> str:
    """Même rendu en PDF, avec la police Poppins incorporée : s'affiche à l'identique partout."""
    import pymupdf

    fonts = {k: pymupdf.Font(fontfile=os.path.join(ASSETS, "fonts", v)) for k, v in FONT_FILES.items()}
    bg_path = os.path.join(ASSETS, "background.jpg")
    doc = pymupdf.open()
    bg_xref = 0
    W, H = 960, 540

    def rgb(hexc):
        return tuple(int(hexc[i:i + 2], 16) / 255 for i in (0, 2, 4))

    for spec in _normalize(slides):
        page = doc.new_page(width=W, height=H)
        if not bg_xref:
            bg_xref = page.insert_image(page.rect, filename=bg_path)
        else:
            page.insert_image(page.rect, xref=bg_xref)
        for c in CARDS.get(spec["mode"], []):
            x, y, w, h = c["rect"]
            shape = page.new_shape()
            shape.draw_rect(pymupdf.Rect(x, y, x + w, y + h), radius=13 / min(w, h))
            shape.finish(fill=rgb(c["fill"]), fill_opacity=c["alpha"], color=rgb(c["line"]) if c.get("line") else None,
                         stroke_opacity=c.get("line_alpha", 1), width=1.5 if c.get("line") else 0)
            shape.commit()
        for op in layout_slide(spec):
            if op["kind"] == "dash":
                page.draw_rect(pymupdf.Rect(op["x"], op["y"], op["x"] + op["w"], op["y"] + op["h"]),
                               color=None, fill=(1, 1, 1))
                continue
            _pdf_text(page, op, fonts, rgb(op["color"]))
    doc.save(out_path, garbage=3, deflate=True)
    return out_path


def _pdf_text(page, op, fonts, color):
    """Écrit un bloc de texte avec retour à la ligne identique au calcul de mise en page."""
    import pymupdf

    size, lh, cs, maxw = op["size"], op["lh"], op["cs"] or 0, op["w"] - 6
    # mots avec leur style et leur opacité
    words = []
    for text, st, *rest in op["runs"]:
        alpha = (rest[0] / 100) if rest else 1.0
        parts = text.split(" ")
        for i, wd in enumerate(parts):
            words.append((wd, st, alpha, i < len(parts) - 1))
    lines, cur, cw = [], [], 0.0
    for wd, st, alpha, space in words:
        w = tw(wd, st, size, cs)
        sp = tw(" ", st, size, cs) if space else 0
        if cur and cw + w > maxw:
            lines.append(cur); cur, cw = [], 0.0
        cur.append((wd, st, alpha, w, sp)); cw += w + sp
    if cur:
        lines.append(cur)
    writers = {}
    for n, line in enumerate(lines):
        baseline = op["y"] + n * lh + lh / 2 + 0.36 * size
        xpos = op["x"]
        for wd, st, alpha, w, sp in line:
            wr = writers.setdefault(alpha, pymupdf.TextWriter(page.rect))
            if cs:
                for ch in wd:
                    wr.append((xpos, baseline), ch, font=fonts[st], fontsize=size)
                    xpos += tw(ch, st, size, cs)
            else:
                wr.append((xpos, baseline), wd, font=fonts[st], fontsize=size)
                xpos += w
            xpos += sp
    for alpha, wr in writers.items():
        wr.write_text(page, color=color, opacity=alpha)
