#!/usr/bin/env python3
"""
PPTX -> PDF Telegram Botu
--------------------------
Kullanıcının gönderdiği .pptx dosyasını, görsel ve metin kalitesini
bozmadan PDF'e dönüştürüp geri gönderir.

Dönüştürme motoru: LibreOffice (soffice --headless).
LibreOffice, PowerPoint'i gerçek bir render motoruyla işlediği için
görselleri yeniden sıkıştırmaz / rasterize etmez; metinler vektörel
kalır, gömülü fontlar ve görseller orijinal çözünürlüğünde kalır.
Bu yüzden "python-pptx ile manuel PDF üretme" gibi kalite kaybına
yol açan yöntemler yerine LibreOffice tercih edildi.

Kurulum:
    pip install python-telegram-bot --upgrade
    sudo apt-get install libreoffice        # soffice komutu için

Çalıştırma:
    export TELEGRAM_BOT_TOKEN="123456:ABC-DEF..."
    python3 bot.py
"""

import asyncio
import logging
import os
import shutil
import subprocess
import tempfile
import uuid

from telegram import Update
from telegram.constants import ChatAction, ParseMode
from telegram.ext import (
    Application,
    ApplicationBuilder,
    ContextTypes,
    MessageHandler,
    CommandHandler,
    filters,
)

# --------------------------------------------------------------------------- #
# Ayarlar
# --------------------------------------------------------------------------- #

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
MAX_FILE_SIZE_MB = 50          # Telegram bot API indirme limiti ~20MB (local API server ile artırılabilir)
SOFFICE_TIMEOUT_SEC = 180       # Büyük dosyalar için dönüştürme zaman aşımı
ALLOWED_EXTENSIONS = (".pptx", ".ppt", ".potx", ".pptm")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("pptx2pdf_bot")


# --------------------------------------------------------------------------- #
# Yardımcı fonksiyonlar
# --------------------------------------------------------------------------- #

def find_soffice() -> str:
    """Sistemde kurulu LibreOffice çalıştırılabilir dosyasını bulur."""
    for candidate in ("soffice", "libreoffice"):
        path = shutil.which(candidate)
        if path:
            return path
    raise RuntimeError(
        "LibreOffice (soffice) bulunamadı. Kurulum için: "
        "sudo apt-get install libreoffice"
    )


def convert_pptx_to_pdf(input_path: str, output_dir: str) -> str:
    """
    LibreOffice'i headless modda çalıştırarak pptx -> pdf dönüştürür.
    Yüksek kaliteli görsel/metin çıktısı için PDF export filtre
    seçenekleri ayarlanır (JPEG sıkıştırması kapalı, çözünürlük yüksek).
    """
    soffice = find_soffice()

    # PDF export filtre seçenekleri:
    # - Quality: JPEG kalitesi (0-100) -> görseller için 100
    # - ReduceImageResolution: false -> görselleri küçültme
    # - UseLosslessCompression: true -> mümkünse kayıpsız sıkıştırma
    filter_options = (
        'impress_pdf_Export:'
        'Quality=100,'
        'ReduceImageResolution=false,'
        'UseLosslessCompression=true,'
        'ExportNotes=false'
    )

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to",
        f"pdf:{filter_options}",
        "--outdir",
        output_dir,
        input_path,
    ]

    logger.info("Dönüştürme komutu çalıştırılıyor: %s", " ".join(cmd))

    # Her dönüştürme için ayrı bir kullanıcı profili (user installation)
    # kullanmak, eşzamanlı isteklerde soffice çakışmalarını önler.
    user_profile_dir = os.path.join(output_dir, "lo_profile")
    env = os.environ.copy()

    cmd_with_profile = cmd + [
        f"-env:UserInstallation=file://{user_profile_dir}"
    ]

    result = subprocess.run(
        cmd_with_profile,
        cwd=output_dir,
        capture_output=True,
        text=True,
        timeout=SOFFICE_TIMEOUT_SEC,
        env=env,
    )

    if result.returncode != 0:
        logger.error("soffice stderr: %s", result.stderr)
        logger.error("soffice stdout: %s", result.stdout)
        raise RuntimeError(f"LibreOffice dönüştürme hatası:\n{result.stderr or result.stdout}")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_dir, base_name + ".pdf")

    if not os.path.exists(pdf_path):
        raise RuntimeError("PDF dosyası oluşturulamadı (beklenmeyen çıktı yolu).")

    return pdf_path


# --------------------------------------------------------------------------- #
# Telegram Handler'ları
# --------------------------------------------------------------------------- #

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Merhaba! 👋\n\n"
        "Bana bir *.pptx* (PowerPoint) dosyası gönder, "
        "sana görsel ve metin kalitesi bozulmadan *PDF* olarak geri göndereyim.\n\n"
        "Not: Dönüştürme LibreOffice ile yapıldığı için orijinal fontlar, "
        "görseller ve düzen olabildiğince korunur.",
        parse_mode=ParseMode.MARKDOWN,
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Kullanım:\n"
        "1) Bir .pptx dosyası gönder (dosya olarak, fotoğraf değil).\n"
        "2) Bot dosyayı PDF'e çevirip sana geri yollar.\n\n"
        f"Maksimum dosya boyutu: {MAX_FILE_SIZE_MB} MB"
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    document = update.message.document
    if document is None:
        return

    file_name = document.file_name or "sunum.pptx"
    ext = os.path.splitext(file_name)[1].lower()

    if ext not in ALLOWED_EXTENSIONS:
        await update.message.reply_text(
            "⚠️ Lütfen bir PowerPoint dosyası gönder (.pptx / .ppt / .pptm / .potx)."
        )
        return

    size_mb = document.file_size / (1024 * 1024) if document.file_size else 0
    if size_mb > MAX_FILE_SIZE_MB:
        await update.message.reply_text(
            f"⚠️ Dosya çok büyük ({size_mb:.1f} MB). "
            f"Maksimum {MAX_FILE_SIZE_MB} MB destekleniyor."
        )
        return

    status_msg = await update.message.reply_text("📥 Dosya indiriliyor...")
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.UPLOAD_DOCUMENT
    )

    work_dir = tempfile.mkdtemp(prefix=f"pptx2pdf_{uuid.uuid4().hex}_")
    try:
        input_path = os.path.join(work_dir, file_name)

        tg_file = await context.bot.get_file(document.file_id)
        await tg_file.download_to_drive(custom_path=input_path)

        await status_msg.edit_text("🔄 PDF'e dönüştürülüyor... (biraz sürebilir)")

        loop = asyncio.get_running_loop()
        pdf_path = await loop.run_in_executor(
            None, convert_pptx_to_pdf, input_path, work_dir
        )

        await status_msg.edit_text("📤 PDF gönderiliyor...")
        await context.bot.send_chat_action(
            chat_id=update.effective_chat.id, action=ChatAction.UPLOAD_DOCUMENT
        )

        pdf_filename = os.path.splitext(file_name)[0] + ".pdf"
        with open(pdf_path, "rb") as pdf_file:
            await update.message.reply_document(
                document=pdf_file,
                filename=pdf_filename,
                caption="✅ Dönüştürme tamamlandı.",
            )

        await status_msg.delete()

    except subprocess.TimeoutExpired:
        logger.exception("Dönüştürme zaman aşımına uğradı")
        await status_msg.edit_text(
            "❌ Dönüştürme zaman aşımına uğradı. Dosya çok büyük veya karmaşık olabilir."
        )
    except Exception as exc:  # noqa: BLE001
        logger.exception("Dönüştürme sırasında hata oluştu")
        await status_msg.edit_text(f"❌ Bir hata oluştu:\n{exc}")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


async def handle_wrong_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Lütfen bana bir *.pptx dosyası* gönder (dosya/document olarak, "
        "fotoğraf ya da metin olarak değil).",
        parse_mode=ParseMode.MARKDOWN,
    )


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error("Beklenmeyen hata:", exc_info=context.error)


# --------------------------------------------------------------------------- #
# Ana giriş noktası
# --------------------------------------------------------------------------- #

def build_app() -> Application:
    if not BOT_TOKEN:
        raise SystemExit(
            "TELEGRAM_BOT_TOKEN ortam değişkeni ayarlı değil.\n"
            'Örn: export TELEGRAM_BOT_TOKEN="123456:ABC-DEF..."'
        )

    application = ApplicationBuilder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    application.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND | filters.PHOTO | filters.VIDEO,
            handle_wrong_type,
        )
    )
    application.add_error_handler(error_handler)

    return application


def main() -> None:
    app = build_app()
    logger.info("Bot başlatılıyor (polling modu)...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
