"""Réception d'un .pptx / .pdf sur Telegram et renvoi au design système.

Branché dans bot.py via register_deck_handler(app).
"""
from __future__ import annotations

import asyncio
import logging
import os
import tempfile
from pathlib import Path

from telegram import Update
from telegram.ext import Application, ContextTypes, MessageHandler, filters

from deckstyle import convert

log = logging.getLogger("deck")
MAX_BYTES = 20 * 1024 * 1024  # limite de téléchargement des bots Telegram


def _allowed(user_id: int) -> bool:
    raw = os.environ.get("ALLOWED_USER_IDS", "").strip()
    if not raw:
        return True
    return str(user_id) in {x.strip() for x in raw.split(",") if x.strip()}


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg or not msg.document:
        return
    if not _allowed(update.effective_user.id):
        await msg.reply_text(f"Accès refusé (ton id : {update.effective_user.id}).")
        return
    doc = msg.document
    name = doc.file_name or "presentation"
    ext = Path(name).suffix.lower()

    if ext in (".key", ".ppt"):
        await msg.reply_text("Ce format n'est pas lu directement. Exportez d'abord le fichier en PowerPoint (.pptx) ou en PDF, puis renvoyez-le.")
        return
    if ext not in (".pptx", ".pdf"):
        await msg.reply_text("Envoyez un fichier .pptx ou .pdf, ou un texte pour générer les visuels de titre.")
        return
    if doc.file_size and doc.file_size > MAX_BYTES:
        await msg.reply_text("Le fichier dépasse 20 Mo, la limite de Telegram pour les bots. Allégez-le (images compressées) ou envoyez-le en PDF.")
        return

    status = await msg.reply_text("Fichier reçu, mise au format en cours… (environ 30 secondes)")
    await context.bot.send_chat_action(msg.chat_id, "upload_document")
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "source" + ext)
        out = os.path.join(tmp, Path(name).stem + "_design.pptx")
        try:
            f = await doc.get_file()
            await f.download_to_drive(src)
            n = await asyncio.to_thread(convert, src, out, log.info)
        except ValueError as e:
            await status.edit_text(str(e))
            return
        except Exception:
            log.exception("Conversion échouée pour %s", name)
            await status.edit_text("La conversion a échoué. Réessayez, ou envoyez le fichier en PDF.")
            return
        with open(out, "rb") as fh:
            await msg.reply_document(fh, filename=os.path.basename(out),
                                     caption=f"{n} slides au format du design système.")
    await status.delete()


def register_deck_handler(app: Application) -> None:
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
