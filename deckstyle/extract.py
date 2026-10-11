"""Extraction du contenu d'un .pptx ou d'un .pdf, slide par slide.

Chaque slide devient un dict :
  {"index": 1, "background": "dark|light|yellow|image|unknown",
   "paragraphs": [{"text": "...", "size": 24.0, "bold": False, "emph": [...], "top": 120.0}]}
`emph` liste les morceaux du paragraphe mis en valeur (gras ou couleur différente).
"""
from __future__ import annotations

import re

import pymupdf
from pptx import Presentation
from pptx.util import Emu


# --------------------------------------------------------------------------- PPTX
def _fill_kind(owner) -> str | None:
    try:
        bg = owner._element.cSld.bg
        if bg is None:
            return None
        fill = owner.background.fill
        if fill.type == 1:  # couleur unie
            return _classify_rgb(fill.fore_color.rgb)
        if fill.type == 6:
            return "image"
    except Exception:
        pass
    return None


def _shape_bg_kind(slide, slide_w, slide_h) -> str:
    # 1) une forme pleine qui couvre la slide (fréquent dans les exports Google Slides)
    for sh in slide.shapes:
        try:
            if (sh.width or 0) >= 0.9 * slide_w and (sh.height or 0) >= 0.9 * slide_h:
                if sh.shape_type == 13:
                    return "image"
                if sh.fill.type == 1:
                    return _classify_rgb(sh.fill.fore_color.rgb)
        except Exception:
            pass
    # 2) fond de la slide, puis de sa mise en page, puis du masque
    for owner in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        k = _fill_kind(owner)
        if k:
            return k
    return "unknown"


def _classify_rgb(rgb) -> str:
    r, g, b = rgb[0], rgb[1], rgb[2]
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    if lum < 90:
        return "dark"
    if r > 200 and g > 180 and b < 170:
        return "yellow"
    return "light"


def _run_color(run):
    try:
        if run.font.color and run.font.color.type is not None and run.font.color.rgb is not None:
            return str(run.font.color.rgb)
    except Exception:
        pass
    return None


def _hex_kind(hexcolor):
    if not hexcolor:
        return None
    return _classify_rgb(bytes.fromhex(hexcolor[-6:]))


def _placeholder_size(sh) -> float:
    try:
        if sh.is_placeholder:
            t = int(sh.placeholder_format.type)
            return {1: 40, 3: 44, 4: 24, 2: 20}.get(t, 18)
    except Exception:
        pass
    return 18


def _iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # groupe
            yield from _iter_shapes(sh.shapes)
        else:
            yield sh


def extract_pptx(path: str) -> list[dict]:
    prs = Presentation(path)
    sw, sh_ = prs.slide_width or 12192000, prs.slide_height or 6858000
    scale = 960 / Emu(sw).pt  # coordonnées ramenées à une slide 960 x 540
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        paras = []
        for shp in _iter_shapes(slide.shapes):
            try:
                w, h = Emu(shp.width or 0).pt * scale, Emu(shp.height or 0).pt * scale
                if 0 < h <= 8 and 15 <= w <= 140 and not (shp.has_text_frame and shp.text_frame.text.strip()):
                    paras.append({"dash": True, "top": round(Emu(shp.top or 0).pt * scale),
                                  "left": round(Emu(shp.left or 0).pt * scale)})
            except Exception:
                pass
        shapes = [s for s in _iter_shapes(slide.shapes) if s.has_text_frame]
        shapes.sort(key=lambda s: ((s.top or 0), (s.left or 0)))
        for sh in shapes:
            for p in sh.text_frame.paragraphs:
                runs = [r for r in p.runs if r.text]
                text = "".join(r.text for r in runs).strip()
                if not text:
                    continue
                sizes = [r.font.size.pt for r in runs if r.font.size]
                if not sizes:  # taille héritée du gabarit : valeur typique selon le type de zone
                    sizes = [_placeholder_size(sh)]
                colors = [_run_color(r) for r in runs]
                weight = {}
                for r, c in zip(runs, colors):
                    weight[c] = weight.get(c, 0) + len(r.text.strip())
                main_color = max(weight, key=weight.get) if weight else None
                emph = [r.text for r, c in zip(runs, colors)
                        if (r.font.bold and not all(x.font.bold for x in runs)) or (c and c != main_color)]
                paras.append({
                    "text": text,
                    "size": round(max(sizes) * scale, 1) if sizes else None,
                    "bold": all(bool(r.font.bold) for r in runs),
                    "italic": all(bool(r.font.italic) for r in runs),
                    "color": main_color,
                    "emph": [e.strip() for e in emph if e.strip()],
                    "runs": [{"t": r.text, "c": c, "b": bool(r.font.bold)} for r, c in zip(runs, colors)],
                    "top": round(Emu(sh.top or 0).pt * scale),
                    "left": round(Emu(sh.left or 0).pt * scale),
                })
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        bg = _shape_bg_kind(slide, sw, sh_)
        if bg in ("unknown", "image"):
            # déduit du texte : texte clair -> fond sombre, texte foncé -> fond clair
            kinds = [_hex_kind(p.get("color")) for p in paras if p.get("color")]
            if kinds:
                main = max(set(kinds), key=kinds.count)
                bg = {"light": "dark", "yellow": "dark", "dark": "light"}.get(main, bg)
        paras.sort(key=lambda p: (p["top"], p["left"]))
        slides.append({"index": i, "background": bg, "paragraphs": paras, "notes": notes})
    return slides


# --------------------------------------------------------------------------- PDF
def _page_bg(page) -> str:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(0.1, 0.1))
    w, h = pix.width, pix.height
    samples = [pix.pixel(x, y) for x, y in ((1, 1), (w - 2, 1), (1, h - 2), (w - 2, h - 2), (w // 2, h - 2))]
    r = sum(s[0] for s in samples) // len(samples)
    g = sum(s[1] for s in samples) // len(samples)
    b = sum(s[2] for s in samples) // len(samples)
    return _classify_rgb((r, g, b))


from .rules import VERSION_RE as _VERSION_RE


def extract_pdf(path: str) -> list[dict]:
    doc = pymupdf.open(path)
    slides = []
    for i, page in enumerate(doc, 1):
        k = 960 / page.rect.width
        d = page.get_text("dict")
        lines = []
        for block in d["blocks"]:
            for line in block.get("lines", []):
                spans = [s for s in line["spans"] if s["text"].strip()]
                if spans:
                    lines.append((line["bbox"], spans))
        # regroupe les lignes proches de même taille en paragraphes
        paras, cur = [], None
        for bbox, spans in sorted(lines, key=lambda l: (round(l[0][1]), l[0][0])):
            size = max(s["size"] for s in spans)
            text = "".join(s["text"] for s in spans)
            colors = [s["color"] for s in spans]
            if (cur and abs(cur["size"] - size) < 0.6 and abs(bbox[0] - cur["left"]) < 4
                    and bbox[1] - cur["_bottom"] < size * 0.9
                    and not _VERSION_RE.search(cur["lines"][-1])  # une référence biblique se termine là
                    and not (size >= 20 and re.search(r"[.!?]$", cur["lines"][-1].strip()))):  # phrases en gros
                cur["text"] = (cur["text"] + " " + text).strip()
                cur["lines"].append(text.strip())
                cur["_spans"] += [{"text": " ", "color": None, "flags": 0, "font": ""}] + list(spans)
                cur["_bottom"] = bbox[3]
            else:
                cur = {"text": text.strip(), "size": round(size, 1), "top": round(bbox[1]), "left": round(bbox[0]),
                       "_bottom": bbox[3], "_spans": list(spans), "lines": [text.strip()]}
                paras.append(cur)
        for p in paras:
            spans = p.pop("_spans"); p.pop("_bottom")
            if len(p["lines"]) == 1:
                p.pop("lines")
            weight = {}
            for sp in spans:  # couleur principale = celle qui couvre le plus de texte
                weight[sp["color"]] = weight.get(sp["color"], 0) + len(sp["text"].strip())
            main = max(weight, key=weight.get)
            bolds = ["bold" in s["font"].lower() or (s["flags"] & 16) for s in spans]
            real = [(sp, b) for sp, b in zip(spans, bolds) if sp["text"].strip()]
            p["bold"] = all(b for _, b in real)
            p["italic"] = all(("italic" in sp["font"].lower() or (sp["flags"] & 2)) for sp, _ in real)
            p["color"] = "%06X" % main
            p["size"] = round(p["size"] * k, 1)
            p["top"], p["left"] = round(p["top"] * k), round(p["left"] * k)
            p["runs"] = [{"t": sp["text"], "c": None if sp["color"] is None else "%06X" % sp["color"], "b": bool(b)}
                         for sp, b in zip(spans, bolds)]
            p["emph"] = [s["text"].strip() for s, b in zip(spans, bolds)
                         if s["text"].strip() and ((b and not p["bold"]) or s["color"] != main)]
        for dr in page.get_drawings():
            r = dr.get("rect")
            if r and 0 < r.height * k <= 8 and 15 <= r.width * k <= 140:
                paras.append({"dash": True, "top": round(r.y0 * k), "left": round(r.x0 * k)})
        for im in page.get_image_info():
            x0, y0, x1, y1 = im["bbox"]
            if 0 < (y1 - y0) * k <= 8 and 15 <= (x1 - x0) * k <= 140:
                paras.append({"dash": True, "top": round(y0 * k), "left": round(x0 * k)})
        paras.sort(key=lambda p: (p["top"], p["left"]))
        slides.append({"index": i, "background": _page_bg(page), "paragraphs": paras, "notes": ""})
    return slides


def extract(path: str) -> list[dict]:
    low = path.lower()
    if low.endswith(".pptx"):
        return extract_pptx(path)
    if low.endswith(".pdf"):
        return extract_pdf(path)
    raise ValueError("Format non pris en charge (attendu : .pptx ou .pdf)")
