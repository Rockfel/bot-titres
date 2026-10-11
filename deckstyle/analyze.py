"""Classe chaque slide dans une ambiance du design système, avec Claude.

Sortie : une liste de slides au format
  {"mode": "cover|plan|statement|reading|question|twocol",
   "blocks": [{"type": ..., "text": ..., "emphasis": [...]}],   # cover / statement / reading / question
   "rows":   [{"num": "I", "text": "...", "ref": "..."}],        # plan
   "cols":   [{"head": "...", "body": "..."}]}                   # twocol
Sans clé API, une analyse par règles simples prend le relais.
"""
from __future__ import annotations

import json
import os
import re

MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
CHUNK = 12  # slides par appel

BLOCK_TYPES = ["eyebrow", "section", "dash", "title", "statement", "intro", "big",
               "ref", "quote", "label", "question", "answer", "footer"]

SCHEMA = {
    "type": "object",
    "properties": {
        "slides": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "index": {"type": "integer"},
                    "mode": {"type": "string", "enum": ["cover", "plan", "statement", "reading", "question", "twocol"]},
                    "blocks": {"type": "array", "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string", "enum": BLOCK_TYPES},
                            "text": {"type": "string"},
                            "emphasis": {"type": "array", "items": {"type": "string"}},
                        },
                        "required": ["type", "text"],
                    }},
                    "rows": {"type": "array", "items": {
                        "type": "object",
                        "properties": {"num": {"type": "string"}, "text": {"type": "string"}, "ref": {"type": "string"}},
                        "required": ["num", "text"],
                    }},
                    "cols": {"type": "array", "items": {
                        "type": "object",
                        "properties": {"head": {"type": "string"}, "body": {"type": "string"}},
                        "required": ["head", "body"],
                    }},
                },
                "required": ["index", "mode"],
            },
        }
    },
    "required": ["slides"],
}

SYSTEM = """Tu reformates des présentations de prédication (église) dans un design système fixe.
On te donne le texte extrait de chaque slide (paragraphes dans l'ordre, taille de police, mots mis en valeur, couleur de fond d'origine).
Pour CHAQUE slide d'entrée, renvoie exactement UNE slide de sortie, dans le même ordre, avec le même "index".

RÈGLE ABSOLUE : recopie le texte mot pour mot. N'invente rien, ne résume rien, ne corrige pas les citations bibliques, ne traduis pas.
Tu peux seulement découper un paragraphe en plusieurs blocs quand il mélange plusieurs rôles (ex. une référence collée à sa citation), et retirer les numéros de slide ou mentions parasites.

Ambiances (mode) :
- cover : 1re slide de titre. Blocs : eyebrow (nom de série), dash, title (titre du message), footer (date / lieu).
- plan : sommaire. Utilise "rows" : num ("—", "I", "II"...), text (intitulé), ref (référence biblique éventuelle, sans le « · »).
- statement : affirmation forte, phrase clé, slide de transition ou de section. Texte blanc sur le fond.
  Blocs possibles : section (intitulé de partie, ex. "I. LE CHARIOT NEUF — 1 CHRONIQUES 13"), dash (petit trait décoratif s'il y en avait un), title (phrase forte courte/moyenne), statement (phrases plus longues ou suite de phrases), intro (petite phrase d'introduction en italique au-dessus), big (2-4 mots très gros, ex. "Il parle."), question + answer (ex. "Qui a parlé ?" / réponse), ref + quote (verset seul en conclusion).
- reading : lecture biblique. Blocs : section (si la slide ouvre une partie), ref (ex. "1 CHRONIQUES 13.4 · LOUIS SEGOND 1910"), quote (citation entre guillemets), label (petite étiquette type "La foule demande"), question/answer.
  Une citation coupée en deux paragraphes donne deux blocs quote consécutifs.
- question : interpellation directe de l'assemblée (souvent un fond jaune ou clair très marqué d'origine). Blocs : dash, statement, question (la phrase-question finale en gras).
- twocol : deux définitions ou deux idées côte à côte. Utilise "cols" (head + body).

Indices : fond sombre d'origine → statement ; fond clair/crème avec références bibliques → reading ; fond jaune → question.
"emphasis" : reprends les morceaux du bloc qui étaient mis en valeur (gras ou couleur d'accent) à l'origine, recopiés exactement tels qu'ils apparaissent dans "text". Si tout le bloc était en valeur, mets le texte entier.
Les références (ref, section, eyebrow) peuvent rester en majuscules comme à l'origine."""


def _call_claude(slides: list[dict]) -> list[dict]:
    import anthropic

    client = anthropic.Anthropic()
    payload = json.dumps(slides, ensure_ascii=False)
    with client.messages.stream(
        model=MODEL,
        max_tokens=16000,
        system=SYSTEM,
        tools=[{"name": "deck", "description": "Slides reformatées", "input_schema": SCHEMA}],
        tool_choice={"type": "tool", "name": "deck"},
        messages=[{"role": "user", "content": f"Slides à reformater :\n{payload}"}],
    ) as stream:
        msg = stream.get_final_message()
    for block in msg.content:
        if block.type == "tool_use":
            return block.input["slides"]
    raise RuntimeError("Réponse inattendue de Claude")


# ------------------------------------------------------------------ règles de secours
REF_RE = re.compile(r"(LOUIS SEGOND|\bSEGOND\b|\b\d+\.\d+(-\d+)?\b.*·|^[1-3]?\s?[A-ZÉÈ][A-ZÉÈ]+ \d+\.\d+)")


def _fallback(slide: dict) -> dict:
    paras = slide["paragraphs"]
    if not paras:
        return {"index": slide["index"], "mode": "statement", "blocks": []}
    has_ref = any(REF_RE.search(p["text"]) and len(p["text"]) < 80 for p in paras)
    mode = "reading" if has_ref else ("question" if slide["background"] == "yellow" else "statement")
    if slide["index"] == 1:
        mode = "cover"
    blocks = []
    big = max((p.get("size") or 0) for p in paras)
    for p in paras:
        t = p["text"]
        if mode == "reading":
            typ = "ref" if (REF_RE.search(t) and len(t) < 80) else "quote"
        elif mode == "cover":
            typ = "title" if (p.get("size") or 0) == big else ("eyebrow" if not blocks else "footer")
        else:
            typ = "title" if len(t) < 120 else "statement"
        blocks.append({"type": typ, "text": t, "emphasis": p.get("emph", [])})
    return {"index": slide["index"], "mode": mode, "blocks": blocks}


def analyze(slides: list[dict], log=print) -> list[dict]:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        log("ANTHROPIC_API_KEY absente : analyse par règles simples")
        return [_fallback(s) for s in slides]
    out = []
    for k in range(0, len(slides), CHUNK):
        part = slides[k:k + CHUNK]
        try:
            res = _call_claude(part)
            by_index = {r.get("index"): r for r in res}
            out += [by_index.get(s["index"]) or _fallback(s) for s in part]
        except Exception as e:  # on ne bloque jamais l'utilisateur
            log(f"Analyse Claude en échec ({e}) : règles simples pour les slides {k + 1}-{k + len(part)}")
            out += [_fallback(s) for s in part]
    return out
