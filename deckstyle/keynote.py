"""Lecture directe des fichiers Apple Keynote (.key), sans passer par un export.

Un .key est une archive zip de fichiers .iwa (protobuf compressés) ; keynote-parser les décode.
On produit le même format que extract.py : une liste de slides avec leurs paragraphes.
"""
from __future__ import annotations

import io
import warnings
import zipfile


def _rgb_hex(c):
    if not c:
        return None
    try:
        return "%02X%02X%02X" % tuple(round(float(c.get(k, 0)) * 255) for k in ("r", "g", "b"))
    except Exception:
        return None


class _Doc:
    def __init__(self, path: str):
        from keynote_parser.codec import IWAFile

        self.objs = {}
        z = zipfile.ZipFile(path)
        if not any(n.endswith(".iwa") for n in z.namelist()):
            inner = next((n for n in z.namelist() if n.endswith("Index.zip")), None)
            if inner:  # ancien format Keynote : Index.zip à l'intérieur
                z = zipfile.ZipFile(io.BytesIO(z.read(inner)))
        with z:
            names = [n for n in z.namelist() if n.endswith(".iwa")]
            if not names:
                raise ValueError("Ce fichier Keynote n'est pas lisible. Exporte-le en PDF ou en PowerPoint.")
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                for n in names:
                    try:
                        d = IWAFile.from_buffer(z.read(n), n).to_dict()
                    except Exception:
                        continue
                    for chunk in d.get("chunks", []):
                        for a in chunk.get("archives", []):
                            if a.get("objects"):
                                self.objs[str(a["header"]["identifier"])] = a["objects"][0]

    def get(self, ref):
        if isinstance(ref, dict):
            ref = ref.get("identifier")
        return self.objs.get(str(ref)) if ref is not None else None

    # propriétés de style en remontant la chaîne des parents
    def style_props(self, ref, key):
        out, seen = {}, set()
        st = self.get(ref)
        chain = []
        while st and id(st) not in seen:
            seen.add(id(st))
            chain.append(st)
            st = self.get((st.get("super") or {}).get("parent"))
        for st in reversed(chain):  # le parent d'abord, l'enfant écrase
            out.update(st.get(key) or {})
        return out


def _geometry(obj):
    s = obj
    while isinstance(s, dict):
        if "geometry" in s:
            return s["geometry"]
        s = s.get("super")
    return None


def _style_at(entries, i):
    cur = None
    for e in entries:
        if e.get("characterIndex", 0) <= i:
            cur = e.get("object")
        else:
            break
    return cur


def _text_paragraphs(doc: _Doc, storage, x, y, scale):
    text = "".join(storage.get("text") or [])
    if not text.strip():
        return []
    para_entries = (storage.get("tableParaStyle") or {}).get("entries", [])
    char_entries = (storage.get("tableCharStyle") or {}).get("entries", [])
    # coupures de paragraphe Keynote : \n,   (paragraphe),   (saut de ligne)
    paras, start = [], 0
    for i, ch in enumerate(text + "\n"):
        if ch in "\n ":
            paras.append((start, i))
            start = i + 1
    out, offset = [], 0.0
    for n, (a, b) in enumerate(paras):
        seg = text[a:b].replace(" ", " ")
        if not seg.strip():
            continue
        pstyle = doc.style_props(_style_at(para_entries, a), "charProperties")
        # découpe en runs aux changements de style de caractère
        cuts = sorted({a, b} | {e.get("characterIndex", 0) for e in char_entries if a < e.get("characterIndex", 0) < b})
        runs = []
        for c0, c1 in zip(cuts, cuts[1:]):
            props = dict(pstyle)
            cref = _style_at(char_entries, c0)
            if cref:
                props.update(doc.style_props(cref, "charProperties"))
            font = (props.get("fontName") or "").lower()
            runs.append({
                "t": text[c0:c1].replace(" ", " "),
                "c": _rgb_hex(props.get("fontColor")),
                "b": bool(props.get("bold")) or any(w in font for w in ("bold", "black", "heavy", "semibold")),
                "i": bool(props.get("italic")) or "italic" in font,
                "size": props.get("fontSize"),
            })
        runs = [r for r in runs if r["t"]]
        if not runs:
            continue
        sizes = [r["size"] for r in runs if r["size"]]
        weight = {}
        for r in runs:
            weight[r["c"]] = weight.get(r["c"], 0) + len(r["t"].strip())
        main = max(weight, key=weight.get)
        real = [r for r in runs if r["t"].strip()]
        all_bold = all(r["b"] for r in real)
        para_top = offset
        offset += (max(sizes) if sizes else 18) * 1.3
        out.append({
            "text": seg.strip(),
            "size": round(max(sizes) * scale, 1) if sizes else 18,
            "bold": all_bold,
            "italic": all(r["i"] for r in real),
            "color": main,
            "emph": [r["t"].strip() for r in real if (r["b"] and not all_bold) or r["c"] != main],
            "runs": [{"t": r["t"], "c": r["c"], "b": r["b"]} for r in runs],
            "top": round((y + para_top) * scale),
            "left": round(x * scale),
        })
    return out


def _walk(doc: _Doc, obj, scale, out, dx=0.0, dy=0.0):
    t = obj.get("_pbtype", "")
    g = _geometry(obj) or {}
    pos = g.get("position") or {}
    size = g.get("size") or {}
    x, y = float(pos.get("x", 0)) + dx, float(pos.get("y", 0)) + dy
    w, h = float(size.get("width", 0)), float(size.get("height", 0))
    if t == "TSD.GroupArchive":
        for ch in obj.get("children", []):
            c = doc.get(ch)
            if c:
                _walk(doc, c, scale, out, x, y)
        return
    storage = doc.get(obj.get("ownedStorage"))
    if storage and storage.get("_pbtype") == "TSWP.StorageArchive":
        out.extend(_text_paragraphs(doc, storage, x, y, scale))
        return
    # petit trait décoratif (image ou forme fine)
    if t in ("TSD.ImageArchive", "TSWP.ShapeInfoArchive", "TSD.ShapeArchive") and 0 < h * scale <= 8 and 15 <= w * scale <= 140:
        out.append({"dash": True, "top": round(y * scale), "left": round(x * scale)})


def _bg_kind(doc: _Doc, slide) -> str:
    from .extract import _classify_rgb

    st = doc.style_props(slide.get("style"), "slideProperties")
    fill = (st.get("fill") or {})
    if "color" in fill:
        hx = _rgb_hex(fill["color"])
        if hx:
            return _classify_rgb(bytes.fromhex(hx))
    if fill:
        return "image"
    return "unknown"


def extract_key(path: str) -> list[dict]:
    doc = _Doc(path)
    show = next((o for o in doc.objs.values() if o.get("_pbtype") == "KN.ShowArchive"), None)
    if not show:
        raise ValueError("Ce fichier Keynote n'est pas lisible. Exporte-le en PDF ou en PowerPoint.")
    width = float((show.get("size") or {}).get("width") or 960)
    scale = 960 / width
    slides = []
    for ref in (show.get("slideTree") or {}).get("slides", []):
        node = doc.get(ref)
        if not node or node.get("isSkipped"):
            continue
        slide = doc.get(node.get("slide"))
        if not slide:
            continue
        paras = []
        drawables = list(slide.get("drawablesZOrder") or slide.get("ownedDrawables") or [])
        # zones titre / corps du gabarit, si elles sont affichées
        for key, flag in (("titlePlaceholder", "titlePlaceholderVisible"), ("bodyPlaceholder", "bodyPlaceholderVisible")):
            if slide.get(key) and slide.get(flag, False):
                drawables.append(slide[key])
        for d in drawables:
            o = doc.get(d)
            if o:
                _walk(doc, o, scale, paras)
        notes = ""
        note = doc.get(slide.get("note"))
        if note:
            ns = doc.get(note.get("containedStorage"))
            if ns:
                notes = "".join(ns.get("text") or []).strip()
        paras.sort(key=lambda p: (p["top"], p["left"]))
        slides.append({"index": len(slides) + 1, "background": _bg_kind(doc, slide), "paragraphs": paras, "notes": notes})
    return slides
