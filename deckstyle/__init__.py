"""Reformatage de présentations (.pptx / .pdf) au design système des prédications."""
from .extract import extract
from .analyze import analyze
from .render import build


def convert(in_path: str, out_path: str, log=print) -> int:
    """Convertit `in_path` (.pptx ou .pdf) et écrit `out_path`. Renvoie le nombre de slides."""
    slides = extract(in_path)
    if not slides:
        raise ValueError("Aucune slide trouvée dans le fichier")
    specs = analyze(slides, log=log)
    build(specs, out_path)
    return len(specs)
