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
from pypdf import PdfWriter
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

    if

