"""Reformatage de présentations (.pptx / .pdf) au design système des prédications."""
from .extract import extract
from .analyze import analyze
from .render import build, build_pdf


def convert(in_path: str, out_path: str, log=print, pdf_path: str | None = None) -> int:
    """Convertit `in_path` (.pptx, .key ou .pdf) en `out_path` (.pptx) et, si demandé, en `pdf_path`
    (même rendu, police Poppins incorporée). Renvoie le nombre de slides."""
    slides = extract(in_path)
    if not slides:
        raise ValueError("Aucune slide trouvée dans le fichier.")
    if not any(p.get("text") for s in slides for p in s["paragraphs"]):
        raise ValueError("Je ne trouve pas de texte dans ce fichier (PDF scanné ou slides en images ?). "
                         "Envoie plutôt le .pptx d'origine.")
    specs = analyze(slides, log=log)
    build(specs, out_path)
    if pdf_path:
        build_pdf(specs, pdf_path)
    return len(specs)
