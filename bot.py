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
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading
import uuid
import zipfile

import gdown
import requests
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt
from pypdf import PdfReader, PdfWriter
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
MAX_FILE_SIZE_MB = 500          # Not: Telegram'ın standart Bot API'si, bota
                                 # DOĞRUDAN yüklenen dosyalarda indirme
                                 # sınırını 20 MB ile kısıtlar (bizim
                                 # kontrolümüz dışında bir Telegram kuralı).
                                 # Bu sınır Google Drive linki ile gönderilen
                                 # dosyalar için geçerli değildir; gdown
                                 # dosyayı doğrudan Drive'dan indirir.
SOFFICE_TIMEOUT_SEC = 600       # Büyük dosyalar için dönüştürme zaman aşımı (10 dk)

# Büyük/ağır sunumları LibreOffice'e tek seferde vermek yerine, aşağıdaki
# eşiklerden biri aşıldığında slayt gruplarına bölüp ayrı ayrı dönüştürüp
# sonra tek PDF'te birleştiriyoruz. Bu, tek bir dev dosyanın LibreOffice'in
# bellek/CPU sınırlarını (özellikle Railway gibi kısıtlı sunucularda)
# zorlamasını ve zaman aşımına takılmasını önlemeye yardımcı olur.
CHUNK_SLIDE_THRESHOLD = 40      # bu slayt sayısından fazlaysa böl
CHUNK_FILE_SIZE_MB_THRESHOLD = 30  # bu boyuttan büyükse de böl (ağır medya)
CHUNK_SIZE = 15                 # her parçada kaç slayt olacak

# Telegram Bot API'sinin bir botun GÖNDEREBİLECEĞİ dosya boyutu sınırı
# 50 MB'dır (botun bir dosyayı İNDİREBİLECEĞİ 20 MB sınırından farklı
# ve ayrı bir kısıtlamadır). Kayıpsız kalite için görselleri koruduğumuzdan
# büyük/görsel ağırlıklı sunumlardan çıkan PDF bu sınırı aşabilir; bu
# durumda PDF'i parçalara bölüp ayrı ayrı göndeririz.
MAX_TELEGRAM_SEND_BYTES = 45 * 1024 * 1024  # 45MB (50MB'a güvenlik payı)
ALLOWED_EXTENSIONS = (".pptx", ".ppt", ".potx", ".pptm")

# Google Drive dosya/paylaşım linklerini yakalamak için desen
DRIVE_URL_PATTERN = re.compile(r"https?://(drive|docs)\.google\.com/\S+", re.IGNORECASE)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("pptx2pdf_bot")


# --------------------------------------------------------------------------- #
# Genel amaçlı font tespiti ve otomatik kurulum
# --------------------------------------------------------------------------- #
# Amaç: Sadece bilinen tek bir fontu değil, ileride gelecek herhangi bir
# sunumdaki eksik fontu da mümkün olduğunca otomatik çözmek. Akış:
#   1) Dosyada kullanılan tüm font isimleri (tema + çalıştırma/run seviyesi
#      + tablo hücreleri) çıkarılır.
#   2) Sistemde zaten kurulu olanlar atlanır.
#   3) Bilinen bir açık kaynak karşılığı varsa (FONT_ALIASES) o zaten
#      build sırasında kurulu olduğu için ekstra işlem gerekmez.
#   4) Kalanlar için Google Fonts'un herkese açık CSS API'si üzerinden
#      aynı isimde bir font aranır ve varsa indirilip sisteme kurulur.
#      Bu, GitHub deposundaki dosya adlandırma biçimini (bazı fontlar artık
#      "variable font" tek dosya olarak dağıtıldığı için) tahmin etmeye
#      çalışmaktan çok daha güvenilirdir.
# Bulunamayan (Google Fonts kataloğunda da olmayan, tamamen özel/lisanslı)
# fontlar için LibreOffice'in varsayılan ikamesi kullanılmaya devam eder;
# bu durumda kesin bir garanti verilemez.

FONT_ALIASES = {
    # Yaygın Microsoft fontları için, Google Fonts'ta bulunmayan ama
    # açık kaynaklı karşılığı build sırasında zaten kurulmuş olanlar.
    # (Karşılıklar Dockerfile'da fontconfig ile eşleştirilir.)
    "tw cen mt": "Poppins",
    "calibri": "Carlito",
    "cambria": "Caladea",
}

DYNAMIC_FONT_DIR = "/usr/share/fonts/truetype/dynamic"
_GOOGLE_FONTS_CSS_URL = "https://fonts.googleapis.com/css2?family={family}"
# Eski bir tarayıcı User-Agent'ı göndermek, Google'ın woff2 yerine
# doğrudan .ttf font dosyası linki döndürmesini sağlar (LibreOffice/
# fontconfig woff2'yi güvenilir şekilde desteklemez).
_OLD_BROWSER_UA = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"

_dynamic_font_attempted = set()   # bu süreç ömrü boyunca denenen fontlar
_dynamic_font_lock = threading.Lock()


def _iter_text_frames(shapes):
    """
    Bir slayttaki tüm metin çerçevelerini dolaşır: normal metin kutuları,
    gruplanmış şekillerin içindekiler (iç içe olsa bile) ve tablo
    hücrelerindeki metinler dahil.
    """
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_text_frames(shape.shapes)
        elif getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame
        elif getattr(shape, "has_text_frame", False):
            yield shape.text_frame


def _extract_theme_fonts(pptx_path: str) -> set:
    """Pptx içindeki tema dosyalarından (theme1.xml, theme2.xml, ...)
    ana/gövde font isimlerini çıkarır."""
    fonts = set()
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            theme_files = [
                n for n in z.namelist()
                if re.match(r"ppt/theme/theme\d+\.xml$", n)
            ]
            for tf in theme_files:
                root = etree.fromstring(z.read(tf))
                for tag in ("majorFont", "minorFont"):
                    el = root.find(f".//a:fontScheme/a:{tag}/a:latin", ns)
                    if el is not None:
                        typeface = el.get("typeface")
                        if typeface and not typeface.startswith("+"):
                            fonts.add(typeface)
    except Exception:  # noqa: BLE001
        logger.exception("Tema fontları okunurken hata oluştu")
    return fonts


def extract_fonts_used(pptx_path: str) -> set:
    """Bir pptx dosyasında (tema dahil) kullanılan tüm font isimlerini döner."""
    fonts = set()
    try:
        prs = Presentation(pptx_path)
        for slide in prs.slides:
            for text_frame in _iter_text_frames(slide.shapes):
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        name = run.font.name
                        if name and not name.startswith("+"):
                            fonts.add(name)
    except Exception:  # noqa: BLE001
        logger.exception("Font taraması sırasında hata oluştu")

    fonts |= _extract_theme_fonts(pptx_path)
    return {f.strip() for f in fonts if f and f.strip()}


def _get_installed_font_families() -> set:
    """fc-list ile sistemde kurulu olan tüm font ailesi isimlerini
    (küçük harfe çevrilmiş) döner."""
    try:
        result = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True, text=True, timeout=15,
        )
        families = set()
        for line in result.stdout.splitlines():
            for name in line.split(","):
                cleaned = name.strip().lower()
                if cleaned:
                    families.add(cleaned)
        return families
    except Exception:  # noqa: BLE001
        logger.exception("fc-list çalıştırılamadı")
        return set()


def _download_google_font(family_name: str) -> bool:
    """
    Google Fonts'un herkese açık CSS API'sinden verilen isimde bir font
    aramayı ve indirmeyi dener. Bulunup indirilirse True döner.
    """
    try:
        family_param = requests.utils.quote(family_name.strip())
        url = _GOOGLE_FONTS_CSS_URL.format(family=family_param)
        css_resp = requests.get(
            url, headers={"User-Agent": _OLD_BROWSER_UA}, timeout=8
        )
        if css_resp.status_code != 200:
            return False

        match = re.search(
            r"url\((https://fonts\.gstatic\.com/[^)]+?\.ttf)\)", css_resp.text
        )
        if not match:
            return False

        font_resp = requests.get(match.group(1), timeout=15)
        if font_resp.status_code != 200 or len(font_resp.content) < 1000:
            return False

        os.makedirs(DYNAMIC_FONT_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9]", "", family_name)
        out_path = os.path.join(DYNAMIC_FONT_DIR, f"{safe_name}-Regular.ttf")
        with open(out_path, "wb") as fh:
            fh.write(font_resp.content)
        return True

    except Exception:  # noqa: BLE001
        logger.exception("Google Fonts'tan '%s' indirilirken hata oluştu", family_name)
        return False


def ensure_fonts_available(pptx_path: str) -> None:
    """
    Dosyada kullanılan fontlardan sistemde kurulu olmayanları tespit eder
    ve mümkünse otomatik olarak temin eder (bilinen ikame veya Google
    Fonts'tan canlı indirme). Bulunamayan fontlar için LibreOffice'in
    varsayılan ikamesi kullanılmaya devam eder — bu fonksiyon en iyi
    çabayı gösterir, %100 garanti vermez.
    """
    used_fonts = extract_fonts_used(pptx_path)
    if not used_fonts:
        return

    installed = _get_installed_font_families()
    downloaded_any = False

    with _dynamic_font_lock:
        for font_name in used_fonts:
            key = font_name.lower()

            if key in installed:
                continue
            if key in FONT_ALIASES:
                # Build sırasında zaten kurulu bilinen bir ikamesi var
                # (Dockerfile'daki fontconfig eşleştirmesi devreye girer).
                continue
            if key in _dynamic_font_attempted:
                continue

            _dynamic_font_attempted.add(key)  # tekrar denemeyi engelle

            if _download_google_font(font_name):
                logger.info("Font otomatik indirildi: %s", font_name)
                downloaded_any = True

    if downloaded_any:
        try:
            subprocess.run(
                ["fc-cache", "-f"], capture_output=True, timeout=30
            )
        except Exception:  # noqa: BLE001
            logger.exception("fc-cache yenilenemedi")


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
# Büyük dosyalar için: parçalama, ayrı ayrı dönüştürme, birleştirme
# --------------------------------------------------------------------------- #

def _keep_only_slides(prs: Presentation, keep_indices: set) -> None:
    """
    Verilen Presentation nesnesinde SADECE keep_indices'teki (0-tabanlı)
    slaytları bırakır, diğerlerini slayt listesinden çıkarır.

    Not: Bu, slaytı sunumun "sıra listesinden" (sldIdLst) kaldırır; alttaki
    slayt XML parçası paket içinde kalır ama artık hiçbir yerden
    referans edilmez. OOXML okuyucuları (LibreOffice dahil) sadece bu
    listeden erişilebilen slaytları işler, bu yüzden dönüştürme sırasında
    çıkarılan slaytlar render edilmez / işlem yüküne dahil olmaz.
    """
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for idx in reversed(range(len(slides))):
        if idx not in keep_indices:
            xml_slides.remove(slides[idx])


def split_pptx_into_chunks(input_path: str, work_dir: str, chunk_size: int) -> list:
    """
    Bir pptx dosyasını, her biri en fazla chunk_size slayt içeren ayrı
    pptx dosyalarına böler. Bölmeye gerek yoksa (slayt sayısı zaten
    küçükse) tek elemanlı [input_path] listesi döner.
    """
    prs_full = Presentation(input_path)
    total = len(prs_full.slides)

    if total <= chunk_size:
        return [input_path]

    chunk_paths = []
    num_chunks = math.ceil(total / chunk_size)

    for c in range(num_chunks):
        start = c * chunk_size
        end = min(start + chunk_size, total)
        keep = set(range(start, end))

        # Her parça için orijinal dosyanın taze bir kopyasını aç, böylece
        # önceki parçalarda yapılan silmeler birbirini etkilemez.
        prs_chunk = Presentation(input_path)
        _keep_only_slides(prs_chunk, keep)

        chunk_path = os.path.join(work_dir, f"chunk_{c:03d}.pptx")
        prs_chunk.save(chunk_path)
        chunk_paths.append(chunk_path)

    return chunk_paths


def split_pdf_by_size(
    pdf_path: str, work_dir: str, max_bytes: int = MAX_TELEGRAM_SEND_BYTES
) -> list:
    """
    Bir PDF dosyasını, Telegram'ın gönderme sınırının altında kalacak
    şekilde sayfa bazında parçalara böler. Boyutun sayfalar arasında
    kabaca eşit dağıldığı varsayılır (görsel yoğunluğu sayfadan sayfaya
    değişebileceği için kesin bir garanti değildir, ama pratikte iyi
    sonuç verir). Bölmeye gerek yoksa tek elemanlı [pdf_path] döner.
    """
    size = os.path.getsize(pdf_path)
    if size <= max_bytes:
        return [pdf_path]

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    if total_pages <= 1:
        return [pdf_path]  # tek sayfa varsa bölecek bir şey yok

    num_parts = math.ceil(size / max_bytes)
    pages_per_part = max(1, math.ceil(total_pages / num_parts))

    part_paths = []
    for start in range(0, total_pages, pages_per_part):
        writer = PdfWriter()
        for p in range(start, min(start + pages_per_part, total_pages)):
            writer.add_page(reader.pages[p])

        part_num = len(part_paths) + 1
        part_path = os.path.join(work_dir, f"part_{part_num:02d}.pdf")
        with open(part_path, "wb") as fh:
            writer.write(fh)
        part_paths.append(part_path)

    return part_paths


def merge_pdfs(pdf_paths: list, output_path: str) -> None:
    """Verilen PDF dosyalarını sırasıyla tek bir PDF'te birleştirir."""
    writer = PdfWriter()
    try:
        for pdf_path in pdf_paths:
            writer.append(pdf_path)
        with open(output_path, "wb") as fh:
            writer.write(fh)
    finally:
        writer.close()


def convert_pptx_to_pdf_chunked(
    input_path: str,
    work_dir: str,
    chunk_size: int = CHUNK_SIZE,
    progress_callback=None,
) -> str:
    """
    Büyük/ağır bir sunumu slayt gruplarına böler, her grubu ayrı ayrı
    PDF'e çevirir ve sonunda hepsini tek bir PDF'te birleştirir.

    progress_callback(done, total) verilirse her parça tamamlandığında
    çağrılır (ilerleme durumu göstermek için).
    """
    chunk_paths = split_pptx_into_chunks(input_path, work_dir, chunk_size)

    if len(chunk_paths) == 1:
        # Bölmeye gerek yoktu, normal (bölünmemiş) yoldan devam et.
        return convert_pptx_to_pdf(chunk_paths[0], work_dir)

    pdf_chunk_paths = []
    total_chunks = len(chunk_paths)

    for i, chunk_path in enumerate(chunk_paths):
        chunk_out_dir = os.path.join(work_dir, f"chunk_out_{i:03d}")
        os.makedirs(chunk_out_dir, exist_ok=True)
        pdf_path = convert_pptx_to_pdf(chunk_path, chunk_out_dir)
        pdf_chunk_paths.append(pdf_path)

        if progress_callback is not None:
            progress_callback(i + 1, total_chunks)

    merged_path = os.path.join(work_dir, "merged.pdf")
    merge_pdfs(pdf_chunk_paths, merged_path)
    return merged_path


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
        for text_frame in _iter_text_frames(slide.shapes):
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
        "görseller ve düzen olabildiğince korunur.\n\n"
        "📏 *20 MB'dan büyük* dosyalar için Drive linki kullan — "
        "doğrudan dosya yüklemede Telegram'ın 20 MB sınırı var, "
        "Drive linkinde bu sınır yok (500 MB'a kadar destekleniyor).\n\n"
        "📚 Çok slaytlı/ağır sunumlarda, işlemi hafifletmek için "
        "dosyayı otomatik olarak parçalara bölüp dönüştürüyor ve "
        "sonunda tek PDF'te birleştiriyorum.",
        parse_mode=ParseMode.MARKDOWN,
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Kullanım:\n"
        "1) Bir .pptx dosyası gönder (dosya olarak, fotoğraf değil).\n"
        "2) Bot dosyayı PDF'e çevirip sana geri yollar.\n\n"
        f"Maksimum dosya boyutu: {MAX_FILE_SIZE_MB} MB"
    )


def _make_progress_callback(loop: asyncio.AbstractEventLoop, status_msg):
    """
    convert_pptx_to_pdf_chunked içinden (ayrı bir thread'den) çağrılacak,
    Telegram durum mesajını ilerleme bilgisiyle güncelleyen bir callback
    üretir. run_coroutine_threadsafe kullanır çünkü status_msg.edit_text
    bir coroutine ve doğrudan başka bir thread'den await edilemez.
    """
    def _callback(done: int, total: int) -> None:
        text = f"🔄 Parçalara bölünerek dönüştürülüyor: {done}/{total} parça tamamlandı..."
        try:
            future = asyncio.run_coroutine_threadsafe(status_msg.edit_text(text), loop)
            future.result(timeout=5)
        except Exception:  # noqa: BLE001
            pass  # İlerleme mesajı güncellenemese de dönüştürme devam etsin

    return _callback


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

        # 1) Dosyada kullanılan fontları tara; sistemde kurulu olmayan
        #    ama Google Fonts'ta bulunabilenleri otomatik indirip kur.
        ext = os.path.splitext(input_path)[1].lower()
        if ext in (".pptx", ".pptm", ".potx"):
            try:
                await loop.run_in_executor(None, ensure_fonts_available, input_path)
            except Exception:  # noqa: BLE001
                logger.exception("Font hazırlığı sırasında hata oluştu, devam ediliyor")

        # 2) Dönüştürmeden önce, PowerPoint'in "otomatik küçült" ile
        #    küçülttüğü ama LibreOffice'in doğru uygulamadığı font
        #    boyutlarını düzelt (metnin taşmasını/çakışmasını önler).
        fixed_path = os.path.join(work_dir, "fixed_" + file_name)
        try:
            was_fixed = await loop.run_in_executor(
                None, fix_autofit_shrink, input_path, fixed_path
            )
        except Exception:  # noqa: BLE001
            logger.exception("Autofit düzeltmesi başarısız, orijinal dosya kullanılacak")
            was_fixed = False

        convert_input_path = fixed_path if was_fixed else input_path

        # 3) Dosya büyük/ağırsa (çok slaytlı ya da büyük boyutlu),
        #    LibreOffice'i tek seferde zorlamak yerine slayt gruplarına
        #    bölüp ayrı ayrı dönüştür, sonra tek PDF'te birleştir.
        file_size_mb = os.path.getsize(convert_input_path) / (1024 * 1024)
        slide_count = 0
        if ext in (".pptx", ".pptm", ".potx"):
            try:
                slide_count = len(Presentation(convert_input_path).slides)
            except Exception:  # noqa: BLE001
                logger.exception("Slayt sayısı okunamadı")

        use_chunking = ext in (".pptx", ".pptm", ".potx") and (
            slide_count > CHUNK_SLIDE_THRESHOLD
            or file_size_mb > CHUNK_FILE_SIZE_MB_THRESHOLD
        )

        if use_chunking:
            await status_msg.edit_text(
                f"🔄 Büyük dosya tespit edildi ({slide_count} slayt, "
                f"{file_size_mb:.0f} MB).\nParçalara bölünüp dönüştürülecek..."
            )
            progress_cb = _make_progress_callback(loop, status_msg)
            pdf_path = await loop.run_in_executor(
                None,
                convert_pptx_to_pdf_chunked,
                convert_input_path,
                work_dir,
                CHUNK_SIZE,
                progress_cb,
            )
        else:
            await status_msg.edit_text("🔄 PDF'e dönüştürülüyor... (biraz sürebilir)")
            pdf_path = await loop.run_in_executor(
                None, convert_pptx_to_pdf, convert_input_path, work_dir
            )

        # PDF, Telegram'ın gönderme sınırını (50MB) aşıyorsa otomatik
        # olarak parçalara böl.
        pdf_parts = await loop.run_in_executor(
            None, split_pdf_by_size, pdf_path, work_dir
        )

        await status_msg.edit_text("📤 PDF gönderiliyor...")
        await context.bot.send_chat_action(
            chat_id=update.effective_chat.id, action=ChatAction.UPLOAD_DOCUMENT
        )

        pdf_filename = os.path.splitext(file_name)[0] + ".pdf"

        if len(pdf_parts) == 1:
            with open(pdf_parts[0], "rb") as pdf_file:
                await update.message.reply_document(
                    document=pdf_file,
                    filename=pdf_filename,
                    caption="✅ Dönüştürme tamamlandı.",
                )
        else:
            base_name = os.path.splitext(pdf_filename)[0]
            total_parts = len(pdf_parts)
            await status_msg.edit_text(
                f"📤 PDF 50MB sınırını aştığı için {total_parts} parçaya "
                f"bölünüp gönderiliyor..."
            )
            for idx, part_path in enumerate(pdf_parts, start=1):
                part_filename = f"{base_name}_bölüm{idx}-{total_parts}.pdf"
                with open(part_path, "rb") as pdf_file:
                    await update.message.reply_document(
                        document=pdf_file,
                        filename=part_filename,
                        caption=(
                            f"✅ Bölüm {idx}/{total_parts} "
                            "(dosya 50MB sınırını aştığı için parçalandı)"
                        ),
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
        error_text = str(exc).lower()
        if "too big" in error_text or "file is too big" in error_text:
            await status_msg.edit_text(
                "⚠️ Bu dosya Telegram'a doğrudan yüklendiği için "
                "*20 MB* sınırına takıldı (bu bir Telegram kısıtlaması, "
                "botun kendi ayarı değil).\n\n"
                "Çözüm: Dosyayı Google Drive'a yükleyip 'Bağlantıya sahip "
                "olan herkes görüntüleyebilir' yaparak linkini bana gönder "
                "— bu yöntemle 500 MB'a kadar dosya kabul edebiliyorum.",
                parse_mode=ParseMode.MARKDOWN,
            )
        else:
            await status_msg.edit_text(f"❌ Dosya indirilemedi:\n{exc}")
        shutil.rmtree(work_dir, ignore_errors=True)
        return

    await convert_and_reply(update, context, input_path, work_dir, status_msg)


def _stream_download_and_validate(url: str, output_path: str, timeout: int = 60):
    """
    Bir URL'den dosyayı belleğe tamamen yüklemeden (akış/stream hâlinde),
    doğrudan diske yazarak indirir. Büyük dosyalarda (yüzlerce MB) bu,
    tüm içeriği tek seferde RAM'e almaktan (requests.get(...).content)
    çok daha güvenlidir — sunucunun belleğini taşırıp botun sessizce
    çökmesini önler.

    İlk parçaya bakarak içeriğin gerçekten bir pptx (zip imzası "PK" ile
    başlar) olup olmadığını kontrol eder; değilse indirmeyi yarıda kesip
    yarım dosyayı siler.

    Dönüş: (başarılı_mı, durum_kodu, content_type) üçlüsü.
    """
    content_type = ""
    try:
        with requests.get(url, timeout=timeout, stream=True) as resp:
            status_code = resp.status_code
            content_type = resp.headers.get("Content-Type", "")

            if status_code != 200:
                return False, status_code, content_type

            looks_valid = None
            with open(output_path, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=1024 * 1024):  # 1 MB'lık parçalar
                    if not chunk:
                        continue
                    if looks_valid is None:
                        looks_valid = (
                            chunk[:2] == b"PK" or "presentationml" in content_type
                        )
                        if not looks_valid:
                            break  # geçersiz içerik (muhtemelen HTML hata sayfası)
                    fh.write(chunk)

            if not looks_valid:
                try:
                    os.remove(output_path)
                except OSError:
                    pass
                return False, status_code, content_type

            return True, status_code, content_type

    except requests.exceptions.RequestException as exc:
        logger.exception("Streaming indirme sırasında ağ hatası")
        return False, None, str(exc)


def _extract_drive_file_id(url: str) -> str:
    """Verilen Drive/Slides linkinden dosya kimliğini (ID) çıkarır."""
    patterns = [
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"/presentation/d/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
        r"/open\?id=([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None


def download_from_drive(drive_url: str, work_dir: str) -> str:
    """
    Bir Google Drive / Google Slides linkinden dosyayı indirir.
    Üç link türünü ayırt eder:

      1) KLASÖR linki (.../drive/folders/ID) -> desteklenmez, çünkü
         içinde birden fazla dosya olabilir. Kullanıcıya net bir hata
         verilir.
      2) Google Slides (canlı doküman) linki (.../presentation/d/ID) ->
         Google'ın herkese açık "export" adresi üzerinden doğrudan
         .pptx olarak indirilir. Ekstra API anahtarı gerekmez.
      3) Normal Drive DOSYA linki (.../file/d/ID/view) -> gdown ile
         indirilir.

    Not: Her durumda dosyanın Drive'da "Bağlantıya sahip olan herkes
    görüntüleyebilir" şeklinde paylaşılmış olması gerekir.
    """
    if "/drive/folders/" in drive_url:
        raise RuntimeError(
            "Bu bir Drive *klasör* linki, tek bir dosya linki değil.\n"
            "Klasördeki .pptx dosyasını aç, sağ tık (ya da ⋮ menüsü) → "
            "'Bağlantı al' ile dosyanın kendi linkini paylaş."
        )

    if "docs.google.com/presentation" in drive_url:
        file_id = _extract_drive_file_id(drive_url)
        if not file_id:
            raise RuntimeError("Google Slides linkinden dosya kimliği okunamadı.")

        # 1. Deneme: "export" adresi (dosya gerçekten native Google Slides
        #    formatındaysa bunu pptx'e dönüştürerek indirir). Bellek
        #    taşmasını önlemek için dosya doğrudan diske akış hâlinde
        #    yazılır (400MB gibi büyük dosyalarda kritik).
        export_url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        candidate_path = os.path.join(work_dir, f"{file_id}.pptx")

        success, status_code, content_type = _stream_download_and_validate(
            export_url, candidate_path, timeout=120
        )

        if success:
            return candidate_path

        # 2. Deneme: Dosya aslında zaten native olmayan (Drive'a yüklenmiş
        #    gerçek bir .pptx dosyası, sadece Slides önizlemesinde açılan)
        #    bir dosya olabilir — bu durumda "export" değil, dosyanın
        #    kendisini doğrudan indirmek gerekir. gdown zaten kendi içinde
        #    diske akış hâlinde yazar (bellek sorunu yaşatmaz).
        try:
            direct_url = f"https://drive.google.com/uc?id={file_id}"
            output_path = gdown.download(
                url=direct_url,
                output=work_dir + os.sep,
                quiet=True,
            )
            if output_path and os.path.exists(output_path):
                return output_path
        except Exception:  # noqa: BLE001
            logger.exception("Doğrudan indirme denemesi de başarısız oldu")

        raise RuntimeError(
            "Google Slides / Drive dosyası indirilemedi.\n"
            f"(Export denemesi -> durum kodu: {status_code}, "
            f"içerik türü: {content_type or 'belirtilmemiş'})\n\n"
            "Dosyanın 'Bağlantıya sahip olan herkes görüntüleyebilir' "
            "şeklinde paylaşıldığından emin ol. Eğer paylaşım ayarı "
            "doğruysa, dosya boyutu/karmaşıklığı Google'ın dışa aktarma "
            "sınırını aşıyor olabilir."
        )

    # Normal Drive dosya linki (.../file/d/ID/view gibi)
    output_path = gdown.download(
        url=drive_url,
        output=work_dir + os.sep,   # sondaki ayraç: orijinal dosya adını kullan
        quiet=True,
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
        # Toplam indirme süresine kesin bir üst sınır koyuyoruz. Google'ın
        # export sunucusu çok büyük/karmaşık dosyalarda bağlantıyı açık
        # tutup uzun süre hiç veri göndermeyebilir; bu durumda tek tek
        # istek zaman aşımları (per-chunk timeout) tetiklenmez ve işlem
        # sonsuza kadar "indiriliyor" durumunda kalabilir. 8 dakikalık
        # bu sınır, o senaryoda botun kilitlenmeden net bir hata
        # vermesini sağlar.
        input_path = await asyncio.wait_for(
            loop.run_in_executor(None, download_from_drive, drive_url, work_dir),
            timeout=480,
        )
    except asyncio.TimeoutError:
        logger.error("Drive indirme işlemi 8 dakikayı aştı, iptal edildi")
        await status_msg.edit_text(
            "⏱️ İndirme 8 dakikadan uzun sürdüğü için iptal edildi.\n\n"
            "Bu genelde dosyanın çok büyük/karmaşık olması ve Google'ın "
            "export işleminin çok uzun sürmesinden kaynaklanır. Öneriler:\n"
            "• Dosyayı gerçek bir .pptx olarak Drive'a yükleyip o linki "
            "paylaş (Slides export'undan çok daha hızlıdır)\n"
            "• Sunumu daha küçük parçalara (örn. 50 slaytlık bölümlere) "
            "ayırıp ayrı ayrı gönder"
        )
        shutil.rmtree(work_dir, ignore_errors=True)
        return
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
