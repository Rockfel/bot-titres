"""
Bot Telegram :
  - envoie un titre        -> reçois les 2 visuels (bandeau + plein écran) ;
  - envoie un .pptx / .pdf -> reçois la présentation refaite au design système (voir deck_handler.py).

Format du message :  Première ligne : deuxième ligne
  - Le texte avant « : » devient la grande ligne, le texte après la ligne secondaire.
  - Sans « : », le titre est réparti automatiquement sur deux lignes.
"""
import io
import os
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from telegram import InputMediaDocument, Update
from telegram.ext import Application, ContextTypes, MessageHandler, filters

from deck_handler import register_deck_handler

BASE = Path(__file__).parent
TOKEN = os.environ["TELEGRAM_TOKEN"]
# Facultatif : ton identifiant Telegram, pour que personne d'autre n'utilise le bot
ALLOWED = {int(x) for x in os.environ.get("ALLOWED_USER_IDS", "").split(",") if x.strip()}

NAVY = (36, 49, 71, 255)
CYAN = (188, 246, 248, 255)
WHITE = (255, 255, 255, 255)


def font(weight, size):
    return ImageFont.truetype(str(BASE / "fonts" / f"Poppins-{weight}.ttf"), size)


def fit(draw, text, weight, size, max_w):
    while size > 10 and draw.textlength(text, font=font(weight, size)) > max_w:
        size -= 1
    return font(weight, size)


def split_title(title):
    title = re.sub(r"\s+", " ", title.replace("'", "’")).strip()
    if ":" in title:
        a, b = title.split(":", 1)
        return a.strip().upper(), b.strip().upper()
    words = title.upper().split()
    if len(words) < 2:
        return title.upper(), ""
    # coupe au plus près du milieu
    best = min(range(1, len(words)),
               key=lambda i: abs(len(" ".join(words[:i])) - len(" ".join(words[i:]))))
    return " ".join(words[:best]), " ".join(words[best:])


def render_bandeau(l1, l2):
    im = Image.open(BASE / "templates" / "bandeau.png").convert("RGBA")
    d = ImageDraw.Draw(im)
    if l2:
        d.text((960, 950), l1, font=fit(d, l1, "Bold", 50, 780), fill=NAVY, anchor="ms")
        d.text((960, 1004), l2, font=fit(d, l2, "Regular", 40, 780), fill=NAVY, anchor="ms")
    else:
        d.text((960, 975), l1, font=fit(d, l1, "Bold", 54, 780), fill=NAVY, anchor="ms")
    return im


def render_plein_ecran(l1, l2):
    im = Image.open(BASE / "templates" / "plein_ecran.png").convert("RGBA")
    d = ImageDraw.Draw(im)
    if l2:
        d.text((960, 420), l1, font=fit(d, l1, "Bold", 130, 1500), fill=WHITE, anchor="ms")
        d.text((960, 580), l2, font=fit(d, l2, "Bold", 110, 1500), fill=CYAN, anchor="ms")
    else:
        d.text((960, 480), l1, font=fit(d, l1, "Bold", 130, 1500), fill=WHITE, anchor="ms")
    return im


def to_file(im, name):
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    buf.name = name
    return buf


async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if ALLOWED and user.id not in ALLOWED:
        await update.message.reply_text(f"Accès refusé (ton id : {user.id}).")
        return
    title = update.message.text.strip()
    if title.startswith("/"):
        await update.message.reply_text(
            "Envoie-moi le titre du message, par ex. :\n"
            "Marchez dans l’amour : l’ordre du Grand Berger\n\n"
            "Ou envoie une présentation (.pptx ou .pdf) : je te la renvoie au format du design système.\n\n"
            f"(ton id Telegram : {user.id})")
        return
    l1, l2 = split_title(title)
    slug = re.sub(r"[^A-Za-z0-9]+", "_", l1)[:40].strip("_") or "titre"
    # envoyés en « document » pour garder la qualité et la transparence du bandeau
    await update.message.reply_media_group([
        InputMediaDocument(to_file(render_bandeau(l1, l2), f"{slug}_bandeau.png")),
        InputMediaDocument(to_file(render_plein_ecran(l1, l2), f"{slug}_plein_ecran.png")),
    ])


def main():
    import asyncio
    asyncio.set_event_loop(asyncio.new_event_loop())  # requis depuis Python 3.14+
    app = Application.builder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT, on_message))
    register_deck_handler(app)
    print("Bot démarré. Ctrl+C pour arrêter.")
    app.run_polling()


if __name__ == "__main__":
    main()
