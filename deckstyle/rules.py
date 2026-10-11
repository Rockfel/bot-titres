"""Reconnaissance des slides par règles (gratuit, sans clé API).

Pensé pour les présentations de prédication : couverture, plan, affirmations sur fond sombre,
lectures bibliques (référence + citation) sur fond clair, questions sur fond jaune.
"""
from __future__ import annotations

import re

VERSIONS = (r"SEGOND|\bLSG\b|\bS21\b|\bNEG\b|\bBDS\b|SEMEUR|DARBY|OSTERVALD|COLOMBE|\bTOB\b|"
            r"JÉRUSALEM|JERUSALEM|PAROLE DE VIE|\bPDV\b|\bNBS\b|KING JAMES|\bKJV\b|\bNIV\b|MARTIN|CRAMPON")
VERSION_RE = re.compile(VERSIONS, re.I)
# « 1 Chroniques 13.4 », « Jean 6.28-29 », « Mt 12:1-8 », « EXODE 18.13 »
REF_RE = re.compile(r"^\s*([1-3]\s?)?[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\.]{1,20}(\s[A-Za-zÀ-ÿ]{2,12})?\s+\d{1,3}\s?[\.,:]\s?\d{1,3}")
ROMAN_RE = re.compile(r"^\s*([IVX]{1,5}|\d{1,2})[\.\)]\s")
NUM_RE = re.compile(r"^\s*([IVX]{1,5}|\d{1,2}|[—–-])\.?\s*$")
OPEN_Q = ("«", "“", '"', "„")
CLOSE_Q = ("»", "”", '"')
SECTION_WORDS = ("INTRODUCTION", "CONCLUSION", "APPEL", "PRIÈRE", "PRIERE")


def _letters(t):
    return [c for c in t if c.isalpha()]


def is_caps(t):
    L = _letters(t)
    return len(L) >= 3 and sum(c.isupper() for c in L) / len(L) > 0.85


def is_ref(p):
    t = p["text"]
    return len(t) < 110 and not t.lstrip().startswith(OPEN_Q) and (bool(VERSION_RE.search(t)) or bool(REF_RE.match(t)) and len(t) < 60)


def is_section(p, max_size):
    t = p["text"]
    if len(t) > 130 or is_ref(p):
        return False
    caps = is_caps(t)
    small = (p.get("size") or 0) <= 0.7 * max_size
    if caps and (small or bool(ROMAN_RE.match(t)) or any(w in t.upper() for w in SECTION_WORDS)):
        return True
    if small and len(t) < 60 and any(w in t.upper() for w in SECTION_WORDS):
        return True
    return small and bool(ROMAN_RE.match(t)) and len(t) < 90


def _dominant_color(paras):
    """Couleur du texte courant de la slide (celle qui couvre le plus de caractères)."""
    weights = {}
    for p in paras:
        for r in p.get("runs") or [{"t": p["text"], "c": p.get("color")}]:
            if r.get("c"):
                weights[r["c"]] = weights.get(r["c"], 0) + len(r["t"].strip())
    return max(weights, key=weights.get) if weights else None


def _emphasis(p, main_color):
    """Morceaux mis en valeur : couleur d'accent ou gras au milieu d'un texte normal."""
    runs = p.get("runs")
    if not runs:
        if main_color and p.get("color") and p["color"] != main_color:
            return [p["text"]]
        return [p["text"]] if p.get("bold") else list(p.get("emph") or [])
    all_bold = all(r["b"] for r in runs if r["t"].strip())
    out, cur, gap = [], "", ""
    for r in runs:
        if not r["t"].strip():  # espace entre deux lignes / deux runs
            gap += r["t"] or " "
            continue
        hot = (main_color and r.get("c") and r["c"] != main_color) or (r["b"] and not all_bold)
        if hot:
            cur = (cur + gap + r["t"]) if cur else r["t"]
        elif cur:
            out.append(cur.strip()); cur = ""
        gap = ""
    if cur:
        out.append(cur.strip())
    if all_bold and not out:
        return [p["text"]]
    return [re.sub(r"\s+", " ", e) for e in out if e]


# ------------------------------------------------------------------ cas particuliers
def _plan(paras):
    items = [p for p in paras[1:] if not p.get("dash")]
    # regroupe par ligne (tops proches), puis de gauche à droite
    items.sort(key=lambda p: p["top"])
    lines, cur = [], []
    for p in items:
        if cur and abs(p["top"] - cur[0]["top"]) > 10:
            lines.append(cur); cur = []
        cur.append(p)
    if cur:
        lines.append(cur)
    items = [p for line in lines for p in sorted(line, key=lambda q: q["left"])]
    rows, pending = [], None
    for p in items:
        t = p["text"].strip()
        if NUM_RE.match(t):
            pending = t.rstrip(".")
            continue
        m = re.match(r"^\s*([IVX]{1,5}|\d{1,2}|[—–-])[\.\)]?\s+(.*)$", t)
        num = pending
        if num is None and m and (m.group(1) in ("—", "–", "-") or ROMAN_RE.match(t) or m.group(1).isdigit()):
            num, t = m.group(1), m.group(2)
        pending = None
        text, ref = t, ""
        if " · " in t:
            text, ref = [x.strip() for x in t.rsplit(" · ", 1)]
        rows.append({"num": num or "—", "text": text, "ref": ref})
    return rows


def _twocol(paras):
    texts = [p for p in paras if not p.get("dash")]
    lefts = sorted({p["left"] for p in texts})
    if len(texts) < 4 or not lefts or lefts[-1] - lefts[0] < 250:
        return None
    split = lefts[0] + (lefts[-1] - lefts[0]) / 2
    cols = [[p for p in texts if p["left"] < split], [p for p in texts if p["left"] >= split]]
    if any(len(c) < 2 for c in cols) or any(is_ref(p) for p in texts):
        return None
    return [{"head": c[0]["text"], "body": " ".join(p["text"] for p in c[1:])} for c in cols]


# ------------------------------------------------------------------ typage des blocs
def _blocks(mode, paras):
    texts = [p for p in paras if not p.get("dash")]
    max_size = max((p.get("size") or 0) for p in texts) or 1
    main_color = _dominant_color(texts)
    blocks, in_quote, k = [], False, 0
    n_text = len(texts)
    bigs = [p for p in texts if len(p["text"].split()) <= 4 and (p.get("size") or 0) >= 0.9 * max_size]

    for i, p in enumerate(paras):
        if p.get("dash"):
            # un trait n'a de sens qu'en tête de bloc, pas en fin de slide
            if any(not q.get("dash") for q in paras[i + 1:]):
                blocks.append({"type": "dash", "text": ""})
            continue
        t = p["text"].strip()
        size = p.get("size") or max_size
        emph = _emphasis(p, main_color)
        nxt = next((q for q in paras[i + 1:] if not q.get("dash")), None)
        k += 1

        if mode == "cover":
            if size >= 0.95 * max_size:
                typ = "title"
            elif not any(b["type"] == "title" for b in blocks):
                typ = "eyebrow"
            else:
                typ = "footer"
            blocks.append({"type": typ, "text": t, "emphasis": [] if typ != "title" else emph})
            continue

        if is_ref(p):
            blocks.append({"type": "ref", "text": t, "emphasis": []}); in_quote = False; continue
        if is_section(p, max_size):
            blocks.append({"type": "section", "text": t, "emphasis": []}); continue

        starts_q = t.startswith(OPEN_Q)
        ends_q = t.rstrip().endswith(CLOSE_Q)
        question_like = t.endswith("?") and len(t) < 80 and not starts_q and not in_quote and nxt is not None \
            and not is_ref(nxt) and not (nxt["text"].startswith(OPEN_Q))
        prev = blocks[-1]["type"] if blocks else None

        if mode == "question":
            last = nxt is None
            typ = "question" if (last and (p.get("bold") or t.endswith("?") or emph)) else "statement"
            blocks.append({"type": typ, "text": t, "emphasis": [] if typ == "question" else emph}); continue

        if mode == "reading":
            label_like = (not starts_q and not in_quote and len(t) < 45 and size < 0.8 * max_size
                          and nxt is not None and nxt["text"].lstrip().startswith(OPEN_Q))
            if prev == "question":
                typ = "answer"
            elif label_like:
                typ = "label"
            elif starts_q or in_quote or prev in ("ref", "label") or prev == "quote" and not question_like:
                typ = "quote"
            elif question_like:
                typ = "question"
            elif size < 0.8 * max_size and len(t) < 45 and nxt is not None:
                typ = "label"
            else:
                typ = "quote"
            if typ == "quote":
                in_quote = (starts_q or in_quote) and not ends_q
            blocks.append({"type": typ, "text": t, "emphasis": emph if typ in ("quote", "answer") else []}); continue

        # statement
        if prev == "question":
            typ = "answer"
        elif prev == "ref" or (starts_q or in_quote) and any(b["type"] == "ref" for b in blocks):
            typ = "quote"
            in_quote = (starts_q or in_quote) and not ends_q
        elif k == 1 and n_text > 1 and (p.get("italic") or size <= 0.7 * max_size) and len(t) < 160:
            typ = "intro"
        elif k == 1 and n_text > 1 and size >= 1.35 * max((q.get("size") or 0) for q in texts[1:]) and len(t) < 120:
            typ = "title"
        elif len(bigs) >= 2 and p in bigs:
            typ = "big"
        elif question_like:
            typ = "question"
        elif n_text - (1 if any(b["type"] == "intro" for b in blocks) else 0) >= 4 or len(t) > 170:
            typ = "statement"
        else:
            typ = "title"
        blocks.append({"type": typ, "text": t, "emphasis": emph if typ not in ("question",) else []})
    return blocks


def classify(slide: dict) -> dict:
    paras = slide["paragraphs"]
    texts = [p for p in paras if not p.get("dash")]
    idx = slide["index"]
    if not texts:
        return {"index": idx, "mode": "statement", "blocks": []}
    first = texts[0]["text"].strip().lower().rstrip(" :")
    if first in ("plan", "sommaire", "au programme", "programme", "les points", "aperçu") and len(texts) >= 3:
        return {"index": idx, "mode": "plan", "rows": _plan(texts)}
    has_ref = any(is_ref(p) for p in texts)
    if idx == 1 and not has_ref:
        return {"index": idx, "mode": "cover", "blocks": _blocks("cover", paras)}
    cols = _twocol(paras)
    if cols:
        return {"index": idx, "mode": "twocol", "cols": cols}
    bg = slide.get("background")
    if bg == "yellow":
        mode = "question"
    elif has_ref:
        quotes = sum(len(p["text"]) for p in texts if not is_ref(p))
        # un verset court seul sur fond sombre reste une affirmation (citation en grand)
        mode = "statement" if (bg == "dark" and quotes < 120 and sum(is_ref(p) for p in texts) == 1) else "reading"
    else:
        mode = "statement"
    return {"index": idx, "mode": mode, "blocks": _blocks(mode, paras)}
