"""Reformatage de présentations (.pptx / .pdf) au design système des prédications."""
from .extract import extract
from .analyze import analyze
from .render import build


def convert(in_path: str, out_path: str, log=print) -> int:
    """Convertit `in_path` (.pptx ou .pdf) et écrit `out_path`. Renvoie le nombre de slides."""
    slides = extract(in_path)
    if not slides:
        raise ValueError("Aucune slide trouvée dans le fichier.")
    if not any(p.get("text") for s in slides for p in s["paragraphs"]):
        raise ValueError("Je ne trouve pas de texte dans ce fichier (PDF scanné ou slides en images ?). "
                         "Envoie plutôt le .pptx d'origine.")
    specs = analyze(slides, log=log)
    build(specs, out_path)
    return len(specs)
