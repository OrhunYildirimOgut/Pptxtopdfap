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
import re
import shutil
import subprocess
import tempfile
import uuid

import gdown
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt
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

# Google Drive dosya/paylaşım linklerini yakalamak için desen
DRIVE_URL_PATTERN = re.compile(r"https?://(drive|docs)\.google\.com/\S+", re.IGNORECASE)

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


def _iter_all_shapes(shapes):
    """Gruplanmış şekiller dahil, bir slayttaki tüm şekilleri (iç içe olanlar
    da dahil) tek tek dolaşır."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_all_shapes(shape.shapes)
        else:
            yield shape


def _get_autofit_scale(text_frame):
    """
    PowerPoint'in bir metin kutusuna uyguladığı "otomatik küçült"
    (shrink text on overflow) oranını pptx XML'inden okur.
    Örn. fontScale="92500" -> 0.925 (yani %92.5'e küçültülmüş) döner.
    Küçültme uygulanmamışsa None döner.
    """
    txBody = text_frame._txBody
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        return None
    norm_autofit = bodyPr.find(qn("a:normAutofit"))
    if norm_autofit is None:
        return None
    font_scale_attr = norm_autofit.get("fontScale")
    if not font_scale_attr:
        return None
    scale = int(font_scale_attr) / 100000.0
    if scale >= 0.999:
        return None
    return scale


def _disable_autofit(text_frame):
    """
    normAutofit etiketini kaldırıp yerine noAutofit ekler; böylece
    biz gerçek (küçültülmüş) font boyutunu yazdıktan sonra LibreOffice
    üzerine bir daha küçültme uygulamaya çalışmaz.
    """
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        return
    norm_autofit = bodyPr.find(qn("a:normAutofit"))
    if norm_autofit is not None:
        bodyPr.remove(norm_autofit)
    if bodyPr.find(qn("a:noAutofit")) is None:
        etree.SubElement(bodyPr, qn("a:noAutofit"))


def fix_autofit_shrink(input_path: str, output_path: str) -> bool:
    """
    PowerPoint'in "metni otomatik küçült" özelliğiyle küçülttüğü ama
    LibreOffice'in PDF'e çevirirken doğru uygulamadığı font boyutlarını,
    gerçek (küçültülmüş) punto değeri olarak dosyaya yazar.

    Böylece LibreOffice, PowerPoint'te ekranda görünenle aynı boyutta
    metin render eder ve metnin görsellerin/diğer öğelerin üzerine
    taşması engellenir.

    Yalnızca OOXML tabanlı formatlar (.pptx, .pptm, .potx) desteklenir;
    eski ikili .ppt formatı python-pptx tarafından okunamadığından
    bu durumda dosya olduğu gibi bırakılır (False döner).

    Dönüş: en az bir düzeltme yapıldıysa True, hiçbir şey
    değiştirilmediyse (veya dosya işlenemediyse) False.
    """
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in (".pptx", ".pptm", ".potx"):
        return False

    try:
        prs = Presentation(input_path)
    except Exception:
        logger.exception("Autofit düzeltmesi için dosya açılamadı, atlanıyor")
        return False

    changed = False

    for slide in prs.slides:
        for shape in _iter_all_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue

            text_frame = shape.text_frame
            scale = _get_autofit_scale(text_frame)
            if scale is None:
                continue

            for paragraph in text_frame.paragraphs:
                if paragraph.font.size is not None:
                    paragraph.font.size = Pt(paragraph.font.size.pt * scale)
                for run in paragraph.runs:
                    if run.font.size is not None:
                        run.font.size = Pt(run.font.size.pt * scale)

            _disable_autofit(text_frame)
            changed = True

    if changed:
        prs.save(output_path)

    return changed


# --------------------------------------------------------------------------- #
# Telegram Handler'ları
# --------------------------------------------------------------------------- #

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Merhaba! 👋\n\n"
        "Bana bir *.pptx* (PowerPoint) dosyası gönder ya da "
        "halka açık bir *Google Drive* linki paylaş, "
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


async def convert_and_reply(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    input_path: str,
    work_dir: str,
    status_msg,
) -> None:
    """
    Verilen pptx dosyasını PDF'e çevirip kullanıcıya geri gönderir.
    Hem Telegram'a doğrudan yüklenen dosyalar hem de Drive'dan indirilen
    dosyalar için ortak dönüştürme/yanıtlama mantığı burada.
    """
    file_name = os.path.basename(input_path)
    try:
        loop = asyncio.get_running_loop()

        # Dönüştürmeden önce, PowerPoint'in "otomatik küçült" ile
        # küçülttüğü ama LibreOffice'in doğru uygulamadığı font
        # boyutlarını düzelt (metnin taşmasını/çakışmasını önler).
        fixed_path = os.path.join(work_dir, "fixed_" + file_name)
        try:
            was_fixed = await loop.run_in_executor(
                None, fix_autofit_shrink, input_path, fixed_path
            )
        except Exception:  # noqa: BLE001
            logger.exception("Autofit düzeltmesi başarısız, orijinal dosya kullanılacak")
            was_fixed = False

        convert_input_path = fixed_path if was_fixed else input_path

        await status_msg.edit_text("🔄 PDF'e dönüştürülüyor... (biraz sürebilir)")

        pdf_path = await loop.run_in_executor(
            None, convert_pptx_to_pdf, convert_input_path, work_dir
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
    input_path = os.path.join(work_dir, file_name)

    try:
        tg_file = await context.bot.get_file(document.file_id)
        await tg_file.download_to_drive(custom_path=input_path)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Telegram dosyası indirilirken hata oluştu")
        await status_msg.edit_text(f"❌ Dosya indirilemedi:\n{exc}")
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    await convert_and_reply(update, context, input_path, work_dir, status_msg)


def download_from_drive(drive_url: str, work_dir: str) -> str:
    """
    Bir Google Drive paylaşım linkinden dosyayı indirir.
    gdown, Drive'ın büyük dosyalarda gösterdiği 'virüs taraması
    yapılamadı' onay adımını otomatik atlar ve orijinal dosya adını
    (uzantısıyla birlikte) content-disposition başlığından alır.

    Not: Dosyanın Drive'da "Bağlantıya sahip olan herkes görüntüleyebilir"
    şeklinde paylaşılmış olması gerekir.
    """
    output_path = gdown.download(
        url=drive_url,
        output=work_dir + os.sep,   # sondaki ayraç: orijinal dosya adını kullan
        quiet=True,
        fuzzy=True,                  # farklı Drive link formatlarını tanır
    )

    if not output_path or not os.path.exists(output_path):
        raise RuntimeError(
            "Drive dosyası indirilemedi. Linkin doğru olduğundan ve "
            "dosyanın 'Bağlantıya sahip olan herkes görüntüleyebilir' "
            "şeklinde paylaşıldığından emin ol."
        )

    return output_path


async def handle_drive_link(
    update: Update, context: ContextTypes.DEFAULT_TYPE, drive_url: str
) -> None:
    status_msg = await update.message.reply_text("🔗 Drive linki alındı, indiriliyor...")
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.UPLOAD_DOCUMENT
    )

    work_dir = tempfile.mkdtemp(prefix=f"pptx2pdf_drive_{uuid.uuid4().hex}_")

    try:
        loop = asyncio.get_running_loop()
        input_path = await loop.run_in_executor(
            None, download_from_drive, drive_url, work_dir
        )
    except Exception as exc:  # noqa: BLE001
        logger.exception("Drive indirme hatası")
        await status_msg.edit_text(f"❌ {exc}")
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        await status_msg.edit_text(
            "⚠️ Drive'daki dosya bir PowerPoint dosyası değil gibi görünüyor "
            f"(bulunan uzantı: {ext or 'yok'}). Lütfen .pptx dosyasına link ver."
        )
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    size_mb = os.path.getsize(input_path) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        await status_msg.edit_text(
            f"⚠️ Dosya çok büyük ({size_mb:.1f} MB). "
            f"Maksimum {MAX_FILE_SIZE_MB} MB destekleniyor."
        )
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    await convert_and_reply(update, context, input_path, work_dir, status_msg)


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = update.message.text or ""
    match = DRIVE_URL_PATTERN.search(text)
    if match:
        await handle_drive_link(update, context, match.group(0))
        return
    await handle_wrong_type(update, context)


async def handle_wrong_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Lütfen bana bir *.pptx dosyası* gönder (dosya/document olarak) "
        "ya da halka açık bir Google Drive linki paylaş.",
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
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text)
    )
    application.add_handler(
        MessageHandler(filters.PHOTO | filters.VIDEO, handle_wrong_type)
    )
    application.add_error_handler(error_handler)

    return application


def main() -> None:
    app = build_app()
    logger.info("Bot başlatılıyor (polling modu)...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
