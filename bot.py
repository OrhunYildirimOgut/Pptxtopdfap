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
import gc
import io
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
from PIL import Image as PILImage
from PIL import ImageOps
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
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

# Akış hâlinde görsel sıkıştırmadan sonra bile bu boyutu aşan sunumları
# daha küçük slayt gruplarıyla dönüştürerek bellek yükünü sınırlıyoruz.
MEMORY_SAFE_PPTX_MB_THRESHOLD = 400

# Otomatik görsel sıkıştırma: dosya bu boyutu aşarsa, PowerPoint'in
# "Resimleri Sıkıştır" özelliğine benzer şekilde görselleri otomatik
# küçültüp yeniden sıkıştırıyoruz. Bu hem Telegram'ın gönderme sınırını
# aşma ihtimalini azaltır hem de LibreOffice'in bellek yükünü düşürür.
IMAGE_COMPRESS_THRESHOLD_MB = 50
IMAGE_MAX_DIMENSION_PX = 1280   # ekranda gösterim için fazlasıyla yeterli
IMAGE_JPEG_QUALITY = 60

# İlk sıkıştırma turundan sonra dosya hâlâ bu boyutun üzerindeyse
# (örn. slaytlarda az sayıda ama aşırı yüksek çözünürlüklü görsel varsa),
# daha da agresif ayarlarla ikinci bir sıkıştırma turu denenir.
IMAGE_SECOND_PASS_THRESHOLD_MB = 80
IMAGE_MAX_DIMENSION_PX_AGGRESSIVE = 900
IMAGE_JPEG_QUALITY_AGGRESSIVE = 45

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
#   1) Dosyada kullanılan tüm font isimleri (slaytlar + düzenler + asıl
#      slaytlar + tema + madde imleri) çıkarılır.
#   2) Sistemde zaten kurulu olanlar atlanır.
#   3) Lisanslı bir Office fontuysa (FONT_ALIASES) ölçülmüş özgür karşılığı
#      kurulur ve fontconfig'e "bunun yerine onu ver" eşleştirmesi yazılır.
#   4) Kalanlar için Google Fonts'un herkese açık CSS API'si üzerinden
#      aynı isimde bir font aranır ve varsa dört stiliyle indirilip kurulur.
#      Bu, GitHub deposundaki dosya adlandırma biçimini (bazı fontlar artık
#      "variable font" tek dosya olarak dağıtıldığı için) tahmin etmeye
#      çalışmaktan çok daha güvenilirdir.
# Bulunamayan (Google Fonts kataloğunda da olmayan, tamamen özel/lisanslı)
# fontlar için LibreOffice'in varsayılan ikamesi kullanılmaya devam eder;
# bu durumda kesin bir garanti verilemez.

FONT_ALIASES = {
    # Microsoft/Office'e ait, lisanslı olduğu için sunucuya kurulamayan ve
    # Google Fonts'ta da bulunmayan fontlar -> yerine kullanılacak özgür font.
    #
    # Karşılıklar GÖRÜNÜŞE göre değil ÖLÇÜME göre seçildi: gerçek fontun ve
    # adayların harf genişlikleri (Türkçe + İngilizce örnek metin üzerinde)
    # karşılaştırıldı. Satırın nereden kırılacağını harf genişliği belirler;
    # genişliği tutmayan bir ikame metni taşırır/kaydırır. Yanlarındaki
    # yüzdeler ortalama genişlik farkıdır (Calibri -> Carlito ölçümü %0,0
    # çıkarak yöntemi doğruladı; eski "Tw Cen MT -> Poppins" ise +%22 idi).
    #
    # Eşleştirme fontconfig'e _write_fontconfig_aliases() ile yazılır; hedef
    # font kurulu değilse Google Fonts'tan indirilir.
    "aptos": "Barlow",                              # -0,6
    "aptos light": "Barlow",
    "aptos semibold": "Barlow",
    "aptos extrabold": "Barlow",
    "aptos black": "Barlow",
    "aptos display": "Sofia Sans Semi Condensed",   # +0,6
    "aptos narrow": "Sofia Sans Semi Condensed",    # -1,0
    "calibri": "Carlito",                           #  0,0 (birebir)
    "calibri light": "Carlito",                     # +1,2
    "cambria": "Caladea",
    "cambria math": "Caladea",
    "tw cen mt": "Carlito",                         # +1,2
    "tw cen mt condensed": "Yanone Kaffeesatz",     # +1,8
    "century gothic": "Inter",                      # +0,1
    "gill sans mt": "Carlito",                      # -1,2
    "gill sans": "Carlito",
    "franklin gothic book": "Gudea",                #  0,0
    "franklin gothic medium": "Quattrocento Sans",  # +0,2
    "garamond": "Crimson Text",                     # +0,1
    "bahnschrift": "Fira Sans",                     # +0,5
    "segoe ui": "Asap",                             #  0,0
    "segoe ui light": "Asap",
    "segoe ui semibold": "Asap",
    "tahoma": "Figtree",                            # +0,1
    "candara": "Source Sans 3",                     # -0,7
    "corbel": "Assistant",                          # -0,2
    "constantia": "Vollkorn",                       # +0,7
    "consolas": "Anonymous Pro",                    # -0,7
    "rockwell": "Bitter",                           # +0,4
    "bookman old style": "Domine",                  # -2,7
    "book antiqua": "Gelasio",                      # -0,5
    "palatino linotype": "Gelasio",                 # -0,5
    "century schoolbook": "Libre Caslon Text",      # +1,4
    "arial narrow": "Archivo Narrow",               # +0,4
}

DYNAMIC_FONT_DIR = "/usr/share/fonts/truetype/dynamic"
_GOOGLE_FONTS_CSS_URL = "https://fonts.googleapis.com/css2"
# Normal + kalın + italik + kalın italik birlikte istenir. Yalnızca normal
# kalınlık indirilirse LibreOffice kalın/italik metni YAPAY olarak üretir;
# yapay kalının genişlikleri gerçeğinden farklıdır ve satırlar başka yerden
# kırılır. Bir ailede istenen stil yoksa Google 400 döner, o yüzden sırayla
# daha az stil isteyen sorgulara düşülür.
_GOOGLE_FONTS_STYLE_LADDER = (
    ":ital,wght@0,400;0,700;1,400;1,700",
    ":wght@400;700",
    "",
)
# DİKKAT: buraya tarayıcı taklidi yapan bir User-Agent KOYMA. Eski bir
# tarayıcı kimliğiyle sorulduğunda Google artık .ttf değil EOT döndürüyor
# (fontconfig okuyamaz); kimliksiz sorguda dört stil de TrueType gelir.
_FONT_FILE_MAGICS = (b"\x00\x01\x00\x00", b"OTTO", b"true")
_FONTCONFIG_ALIAS_FILENAME = "35-pptx2pdf-aliases.conf"

_dynamic_font_attempted = set()   # bu süreç ömrü boyunca denenen fontlar
_dynamic_font_lock = threading.Lock()


_FONT_TAG_PATTERN = re.compile(
    rb"<(?:\w+:)?(?:latin|ea|cs|sym|buFont)\b[^>]*?\btypeface=\"([^\"]+)\""
)
_FONT_XML_PARTS = re.compile(
    r"ppt/(?:slides|slideLayouts|slideMasters|theme|notesMasters)/[^/]+\.xml$"
)


def extract_fonts_used(pptx_path: str) -> set:
    """
    Bir pptx dosyasında kullanılan tüm font isimlerini döner.

    Dosya python-pptx ile AÇILMAZ (o, dosyanın tamamını belleğe alır);
    yalnızca slayt / düzen / asıl slayt / tema XML'leri zip'ten okunup
    taranır. Eskiden yalnızca slaytlardaki metin parçalarına bakılıyordu;
    oysa başlık ve gövde metinlerinin fontu çoğu zaman slaytta değil ASIL
    SLAYTTA (slideMaster) ve düzenlerde tanımlıdır, madde imlerinin fontu
    (buFont) da ayrıdır — bunlar gözden kaçıyordu.

    Temadaki dile özel uzun liste (<a:font script="Jpan" .../> vb.)
    bilerek alınmaz: onlarca Uzak Doğu fontunu tek tek aramaya yol açar.
    """
    fonts = set()
    try:
        with zipfile.ZipFile(pptx_path) as z:
            for name in z.namelist():
                if not _FONT_XML_PARTS.match(name):
                    continue
                for raw in _FONT_TAG_PATTERN.findall(z.read(name)):
                    typeface = raw.decode("utf-8", "replace").strip()
                    if typeface and not typeface.startswith("+"):
                        fonts.add(typeface)
    except Exception:  # noqa: BLE001
        logger.exception("Font taraması sırasında hata oluştu")
    return fonts


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
    Google Fonts'un herkese açık CSS API'sinden verilen isimde bir fontu
    arar ve bulduğu bütün stilleri (normal / kalın / italik / kalın italik)
    indirir. En az bir stil indirildiyse True döner.
    """
    family_name = family_name.strip()
    try:
        css_text = None
        for style_query in _GOOGLE_FONTS_STYLE_LADDER:
            css_resp = requests.get(
                _GOOGLE_FONTS_CSS_URL,
                params={"family": family_name + style_query},
                timeout=8,
            )
            if css_resp.status_code == 200 and "@font-face" in css_resp.text:
                css_text = css_resp.text
                break
        if css_text is None:
            return False

        os.makedirs(DYNAMIC_FONT_DIR, exist_ok=True)
        safe_name = re.sub(r"[^A-Za-z0-9]", "", family_name)
        saved = 0

        for block in re.findall(r"@font-face\s*{[^}]*}", css_text):
            url_match = re.search(r"url\((https://fonts\.gstatic\.com/[^)]+)\)", block)
            if not url_match:
                continue
            weight_match = re.search(r"font-weight:\s*(\d+)", block)
            weight = weight_match.group(1) if weight_match else "400"
            italic = "Italic" if re.search(r"font-style:\s*italic", block) else ""

            font_resp = requests.get(url_match.group(1), timeout=20)
            data = font_resp.content
            if (
                font_resp.status_code != 200
                or len(data) < 1000
                or data[:4] not in _FONT_FILE_MAGICS
            ):
                # TrueType/OpenType değilse (EOT, woff2...) fontconfig
                # okuyamaz; kurulmuş gibi görünüp işe yaramayacağına atla.
                logger.warning(
                    "'%s' için beklenmeyen font biçimi, atlandı (%r)",
                    family_name, data[:4],
                )
                continue

            out_path = os.path.join(
                DYNAMIC_FONT_DIR, f"{safe_name}-{weight}{italic}.ttf"
            )
            with open(out_path, "wb") as fh:
                fh.write(data)
            saved += 1

        return saved > 0

    except Exception:  # noqa: BLE001
        logger.exception("Google Fonts'tan '%s' indirilirken hata oluştu", family_name)
        return False


def _write_fontconfig_aliases() -> None:
    """
    FONT_ALIASES tablosunu fontconfig'e yazar ("X fontu istenirse Y'yi ver").
    Bu dosya olmadan tablo hiçbir işe yaramaz: LibreOffice bulamadığı font
    için kendi varsayılanına (çok daha geniş bir fonta) düşer.

    binding="same", fontconfig'in kendi ölçü-uyumlu eşleştirmelerinde
    (30-metric-aliases.conf) kullandığı bağlamadır; LibreOffice bunu "aynı
    font" sayıp başka bir ikame aramaz. İçerik değişmediyse dosyaya
    dokunulmaz.
    """
    lines = [
        '<?xml version="1.0"?>',
        '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">',
        "<fontconfig>",
    ]
    for source, target in sorted(FONT_ALIASES.items()):
        lines.append(
            f'  <alias binding="same"><family>{source}</family>'
            f"<accept><family>{target}</family></accept></alias>"
        )
    lines.append("</fontconfig>")
    content = "\n".join(lines) + "\n"

    user_conf_dir = os.path.join(
        os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config"),
        "fontconfig", "conf.d",
    )
    for conf_dir in ("/etc/fonts/conf.d", user_conf_dir):
        path = os.path.join(conf_dir, _FONTCONFIG_ALIAS_FILENAME)
        try:
            if os.path.exists(path):
                with open(path, encoding="utf-8") as fh:
                    if fh.read() == content:
                        return
            os.makedirs(conf_dir, exist_ok=True)
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
            return
        except OSError:
            continue  # sistem dizinine yazılamıyorsa kullanıcı dizinini dene
    logger.warning("Fontconfig eşleştirme dosyası yazılamadı")


def ensure_fonts_available(pptx_path: str) -> None:
    """
    Dosyada kullanılan fontlardan sistemde kurulu olmayanları tespit eder
    ve mümkünse otomatik olarak temin eder: aynı isimle Google Fonts'ta
    varsa onu, yoksa FONT_ALIASES'taki ölçülmüş karşılığını indirir.
    Hiçbiri yoksa LibreOffice'in varsayılan ikamesi kullanılır — bu
    fonksiyon en iyi çabayı gösterir, %100 garanti vermez.
    """
    used_fonts = extract_fonts_used(pptx_path)
    if not used_fonts:
        return

    installed = _get_installed_font_families()
    downloaded_any = False

    with _dynamic_font_lock:
        _write_fontconfig_aliases()

        for font_name in used_fonts:
            key = font_name.lower()
            if key in installed:
                continue

            # Lisanslı Office fontu: adıyla aramak boşuna, karşılığını kur.
            wanted = FONT_ALIASES.get(key, font_name)
            wanted_key = wanted.lower()
            if wanted_key in installed or wanted_key in _dynamic_font_attempted:
                continue

            _dynamic_font_attempted.add(wanted_key)  # tekrar denemeyi engelle

            if _download_google_font(wanted):
                installed.add(wanted_key)
                downloaded_any = True
                if wanted == font_name:
                    logger.info("Font otomatik indirildi: %s", font_name)
                else:
                    logger.info("Font karşılığı indirildi: %s -> %s", font_name, wanted)
            else:
                logger.warning(
                    "Font bulunamadı, LibreOffice varsayılanı kullanılacak: %s",
                    font_name,
                )

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

    Metin boyutlarına ve normAutofit ayarlarına dokunulmaz. LibreOffice
    fontScale / lnSpcReduction ile şablondan gelen puntoyu birlikte
    uygular; bu bilgileri silmek metni büyütüp görsellerin altına taşırır.
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

_SHRINK_IMAGE_EXTS = (
    ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp",
)
# PDF statik bir format: gömülü video/ses hiçbir şekilde PDF'e girmez, ama
# büyük sunumların boyutunun çoğu genelde bunlardır. Parçayı paketten
# SİLMEK ilişkileri (rels) bozar; bunun yerine içeriğini boşaltıyoruz —
# slayttaki kapak görseli (poster frame) ayrı bir görsel olduğu için
# PDF'te aynen görünmeye devam eder.
_SHRINK_MEDIA_EXTS = (
    ".mp4", ".m4v", ".mov", ".avi", ".wmv", ".mpg", ".mpeg", ".mkv", ".webm",
    ".mp3", ".wav", ".m4a", ".wma", ".aac", ".ogg",
)
_SHRINK_MIN_IMAGE_BYTES = 150 * 1024  # bundan küçük görselle uğraşmaya değmez


def _recompress_image_blob(blob: bytes, max_dimension: int, jpeg_quality: int):
    """
    Tek bir görselin baytlarını küçültüp yeniden sıkıştırır. Daha küçük bir
    sonuç elde edilemezse (veya görsel açılamazsa) None döner.

    Saydamlığı olan görseller (logo, kesilmiş figür vb.) JPEG'e çevrilirse
    arka planları SİYAH olur; bunlar PNG olarak kalır, sadece küçültülür.
    """
    try:
        pil_img = PILImage.open(io.BytesIO(blob))
        orig_format = pil_img.format
        orig_w, orig_h = pil_img.size
    except Exception:  # noqa: BLE001
        return None

    if orig_format not in ("JPEG", "PNG", "BMP", "TIFF", "GIF", "WEBP"):
        return None

    # JPEG "draft" modu: 40-50MP'lik bir fotoğrafı tam çözünürlükte decode
    # etmeden, baştan küçük boyutta açar — tepe belleği ciddi düşürür.
    if orig_format == "JPEG" and max(orig_w, orig_h) > max_dimension:
        try:
            pil_img.draft("RGB", (max_dimension, max_dimension))
        except Exception:  # noqa: BLE001
            pass

    try:
        pil_img.load()
        # LibreOffice JPEG'i EXIF yönüne göre döndürerek açar; yeniden
        # kaydederken EXIF kaybolduğu için dönüşü piksellere işliyoruz ki
        # görsel PDF'te yan yatmasın.
        pil_img = ImageOps.exif_transpose(pil_img)
    except Exception:  # noqa: BLE001
        return None

    has_alpha = pil_img.mode in ("RGBA", "LA") or (
        pil_img.mode == "P" and "transparency" in pil_img.info
    )

    w, h = pil_img.size
    if max(w, h) > max_dimension:
        scale = max_dimension / max(w, h)
        if has_alpha and pil_img.mode != "RGBA":
            pil_img = pil_img.convert("RGBA")
        elif not has_alpha and pil_img.mode not in ("RGB", "L"):
            pil_img = pil_img.convert("RGB")
        pil_img = pil_img.resize(
            (max(1, int(w * scale)), max(1, int(h * scale))),
            PILImage.LANCZOS,
        )

    buf = io.BytesIO()
    try:
        if has_alpha:
            pil_img.save(buf, format="PNG", optimize=True)
        else:
            if pil_img.mode not in ("RGB", "L"):
                pil_img = pil_img.convert("RGB")
            pil_img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
    except Exception:  # noqa: BLE001
        return None

    new_blob = buf.getvalue()
    return new_blob if len(new_blob) < len(blob) else None


def shrink_pptx_streaming(
    input_path: str,
    output_path: str,
    max_dimension: int = IMAGE_MAX_DIMENSION_PX,
    jpeg_quality: int = IMAGE_JPEG_QUALITY,
    progress_callback=None,
) -> tuple:
    """
    Büyük bir pptx'i, python-pptx ile AÇMADAN (yani dosyanın tamamını
    belleğe almadan) küçültür. pptx bir zip'tir: girdiler tek tek okunur,
    ppt/media altındaki görseller birer birer yeniden sıkıştırılır,
    video/ses içerikleri boşaltılır, geri kalan her şey olduğu gibi akış
    hâlinde kopyalanır. Tepe bellek ≈ o an işlenen TEK görsel; dosyanın
    500MB olması belleği etkilemez.

    Görselin biçimi değişse de (PNG -> JPEG) zip içindeki adı aynı kalır;
    LibreOffice görsel türünü uzantıdan değil içerikten tanır
    (compress_pptx_images de aynı varsayıma dayanıyor).

    Dönüş: (sıkıştırılan_görsel_sayısı, boşaltılan_medya_sayısı).
    """
    compressed_count = 0
    blanked_count = 0

    with zipfile.ZipFile(input_path) as zin:
        infos = zin.infolist()
        media_total = sum(
            1 for i in infos if i.filename.lower().startswith("ppt/media/")
        )
        media_done = 0

        with zipfile.ZipFile(output_path, "w", allowZip64=True) as zout:
            for info in infos:
                name_lower = info.filename.lower()
                ext = os.path.splitext(name_lower)[1]
                is_media = name_lower.startswith("ppt/media/")

                if is_media and ext in _SHRINK_MEDIA_EXTS:
                    zout.writestr(info.filename, b"", zipfile.ZIP_STORED)
                    blanked_count += 1
                elif (
                    is_media
                    and ext in _SHRINK_IMAGE_EXTS
                    and info.file_size >= _SHRINK_MIN_IMAGE_BYTES
                ):
                    blob = zin.read(info)
                    new_blob = _recompress_image_blob(
                        blob, max_dimension, jpeg_quality
                    )
                    if new_blob is not None:
                        compressed_count += 1
                        blob = new_blob
                    # Görseller zaten sıkıştırılmış veri; tekrar deflate
                    # etmek sadece CPU harcar.
                    zout.writestr(info.filename, blob, zipfile.ZIP_STORED)
                    del blob, new_blob
                    gc.collect()
                else:
                    compress_type = (
                        zipfile.ZIP_STORED if is_media else zipfile.ZIP_DEFLATED
                    )
                    out_info = zipfile.ZipInfo(info.filename, info.date_time)
                    out_info.compress_type = compress_type
                    with zin.open(info) as src, zout.open(out_info, "w") as dst:
                        shutil.copyfileobj(src, dst, 1024 * 1024)

                if is_media:
                    media_done += 1
                    if progress_callback is not None and (
                        media_done % 10 == 0 or media_done == media_total
                    ):
                        progress_callback(media_done, media_total)

    return compressed_count, blanked_count


def count_slides_fast(pptx_path: str) -> int:
    """Slayt sayısını, dosyanın tamamını belleğe yüklemeden (yalnızca
    ppt/presentation.xml'i okuyarak) döndürür."""
    with zipfile.ZipFile(pptx_path) as z:
        root = etree.fromstring(z.read("ppt/presentation.xml"))
    sld_id_lst = root.find(qn("p:sldIdLst"))
    return 0 if sld_id_lst is None else len(sld_id_lst)


def _keep_only_slides(prs: Presentation, keep_indices: set) -> None:
    """
    Verilen Presentation nesnesinde SADECE keep_indices'teki (0-tabanlı)
    slaytları bırakır, diğerlerini tamamen kaldırır.

    ÖNEMLİ: Sadece sldIdLst'ten çıkarmak yetmez — altta yatan ilişkiyi
    (rel) de koparmak gerekir (prs.part.drop_rel). Aksi hâlde slayt
    parçası pakette "yetim" olarak kalır; bu dosya tekrar açılıp
    kaydedildiğinde (örn. sıkıştırma adımında) zip içinde çakışan
    parça adları (ör. iki adet "slide1.xml") oluşmasına ve dosyanın
    bozulmasına yol açabilir.
    """
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for idx in reversed(range(len(slides))):
        if idx not in keep_indices:
            slide_id_elem = slides[idx]
            r_id = slide_id_elem.get(qn("r:id"))
            xml_slides.remove(slide_id_elem)
            if r_id:
                try:
                    prs.part.drop_rel(r_id)
                except Exception:  # noqa: BLE001
                    pass


def write_pptx_chunk(
    input_path: str, chunk_path: str, start: int, end: int
) -> None:
    """
    input_path'teki sunumun [start, end) aralığındaki slaytlarını ayrı bir
    pptx olarak chunk_path'e yazar.

    Parçalar TEK TEK üretilir (hepsi baştan değil): aynı anda bellekte
    yalnızca bir Presentation, diskte yalnızca bir parça bulunur.
    """
    prs_chunk = Presentation(input_path)
    _keep_only_slides(prs_chunk, set(range(start, end)))
    prs_chunk.save(chunk_path)
    del prs_chunk
    gc.collect()


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
    already_shrunk: bool = False,
) -> str:
    """
    Büyük/ağır bir sunumu slayt gruplarına böler, her grubu ayrı ayrı
    PDF'e çevirir ve sonunda hepsini tek bir PDF'te birleştirir.

    Her parça sırayla üretilir -> dönüştürülür -> ara dosyaları silinir;
    diskte hiçbir an tek parçadan fazlası (ve biriken küçük PDF'ler
    dışında bir şey) durmaz.

    progress_callback(done, total) verilirse her parça tamamlandığında
    çağrılır (ilerleme durumu göstermek için).

    already_shrunk=True ise görseller shrink_pptx_streaming ile zaten
    küçültülmüştür; parça başına ikinci bir sıkıştırma yapılmaz.
    """
    total = count_slides_fast(input_path)

    if total <= chunk_size:
        # Bölmeye gerek yok, normal (bölünmemiş) yoldan devam et.
        return convert_pptx_to_pdf(input_path, work_dir)

    pdf_chunk_paths = []
    total_chunks = math.ceil(total / chunk_size)

    for i in range(total_chunks):
        chunk_out_dir = os.path.join(work_dir, f"chunk_out_{i:03d}")
        os.makedirs(chunk_out_dir, exist_ok=True)

        chunk_path = os.path.join(chunk_out_dir, f"chunk_{i:03d}.pptx")
        write_pptx_chunk(
            input_path, chunk_path, i * chunk_size, min((i + 1) * chunk_size, total)
        )

        # Her parçayı LibreOffice'e vermeden önce, SADECE O PARÇAYI
        # (tüm dosyayı değil) sıkıştır. Bu, bellek yükünü çok düşük
        # tutar — aynı anda hafızada sadece birkaç slaytlık görsel
        # bulunur, devasa dosyanın tamamı değil.
        convert_source = chunk_path
        if not already_shrunk:
            try:
                compressed_chunk_path = os.path.join(chunk_out_dir, "compressed.pptx")
                count, _saved = compress_pptx_images(
                    chunk_path,
                    compressed_chunk_path,
                    max_dimension=IMAGE_MAX_DIMENSION_PX,
                    jpeg_quality=IMAGE_JPEG_QUALITY,
                )
                if count > 0 and os.path.exists(compressed_chunk_path):
                    convert_source = compressed_chunk_path
            except Exception:  # noqa: BLE001
                logger.exception(
                    "Parça %d sıkıştırması başarısız, orijinal parça kullanılacak", i
                )

        pdf_path = convert_pptx_to_pdf(convert_source, chunk_out_dir)

        # Parçanın PDF'i dışındaki her şeyi (parça pptx'leri, LibreOffice
        # profili) hemen sil; disk kullanımı parça sayısıyla büyümesin.
        for entry in os.listdir(chunk_out_dir):
            entry_path = os.path.join(chunk_out_dir, entry)
            if entry_path == pdf_path:
                continue
            if os.path.isdir(entry_path):
                shutil.rmtree(entry_path, ignore_errors=True)
            else:
                try:
                    os.remove(entry_path)
                except OSError:
                    pass

        pdf_chunk_paths.append(pdf_path)

        if progress_callback is not None:
            progress_callback(i + 1, total_chunks)

    merged_path = os.path.join(work_dir, "merged.pdf")
    merge_pdfs(pdf_chunk_paths, merged_path)
    return merged_path


def _iter_all_shapes(shapes):
    """Gruplanmış şekiller dahil, bir slayttaki tüm şekilleri (iç içe
    olanlar da dahil) tek tek dolaşır."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_all_shapes(shape.shapes)
        else:
            yield shape


def compress_pptx_images(
    input_path: str,
    output_path: str,
    max_dimension: int = IMAGE_MAX_DIMENSION_PX,
    jpeg_quality: int = IMAGE_JPEG_QUALITY,
) -> tuple:
    """
    PowerPoint'in "Resimleri Sıkıştır" özelliğine benzer şekilde, bir
    pptx dosyasındaki tüm görselleri yeniden boyutlandırıp JPEG olarak
    yeniden sıkıştırır. Büyük/yüksek çözünürlüklü görsellerle dolu
    sunumlarda dosya boyutunu ciddi oranda düşürür.

    ÖNEMLİ: Picture.image, pakette saklanan gerçek parçanın DEĞİL,
    tek kullanımlık bir kopyasının döndürür; bu yüzden gerçek
    ImagePart'a shape.part.related_part(rId) ile erişip onun _blob
    özelliğini değiştiriyoruz — aksi hâlde değişiklik kaydedilmez.

    Vektörel formatlar (EMF/WMF gibi) Pillow tarafından açılamadığından
    atlanır, olduğu gibi bırakılır.

    Dönüş: (sıkıştırılan_görsel_sayısı, kazanılan_byte) ikilisi.
    """
    prs = Presentation(input_path)
    seen_ids = set()
    compressed_count = 0
    saved_bytes = 0

    for slide in prs.slides:
        for shape in _iter_all_shapes(slide.shapes):
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue

            rId = shape._pic.blip_rId
            if rId is None:
                continue

            try:
                image_part = shape.part.related_part(rId)
            except Exception:  # noqa: BLE001
                continue

            if id(image_part) in seen_ids:
                continue
            seen_ids.add(id(image_part))

            blob = image_part.blob

            # Önce format/boyut bilgisini AÇIP TAMAMEN DECODE ETMEDEN
            # oku (PIL.Image.open lazy'dir, .load() çağrılana kadar
            # piksel verisini belleğe almaz).
            try:
                pil_img = PILImage.open(io.BytesIO(blob))
                orig_format = pil_img.format
                orig_w, orig_h = pil_img.size
            except Exception:  # noqa: BLE001
                continue

            if orig_format not in ("JPEG", "PNG", "BMP", "TIFF", "GIF", "WEBP"):
                continue

            # JPEG'lerde "draft" modu: görseli TAM ÇÖZÜNÜRLÜKTE decode
            # edip SONRA küçültmek yerine, JPEG'in kendi DCT ölçekleme
            # özelliğini kullanarak DAHA BAŞTAN küçük boyutta decode
            # eder. Özellikle telefon kameralarından gelen çok yüksek
            # çözünürlüklü (40-50MP) fotoğraflarda tepe bellek
            # kullanımını çok ciddi oranda azaltır.
            if orig_format == "JPEG" and max(orig_w, orig_h) > max_dimension:
                try:
                    pil_img.draft("RGB", (max_dimension, max_dimension))
                except Exception:  # noqa: BLE001
                    pass

            try:
                pil_img.load()
            except Exception:  # noqa: BLE001
                continue  # EMF/WMF gibi Pillow'un açamadığı formatlar / bozuk veri

            w, h = pil_img.size
            if max(w, h) > max_dimension:
                scale = max_dimension / max(w, h)
                pil_img = pil_img.resize(
                    (max(1, int(w * scale)), max(1, int(h * scale))),
                    PILImage.LANCZOS,
                )

            if pil_img.mode in ("RGBA", "P", "LA"):
                pil_img = pil_img.convert("RGB")

            buf = io.BytesIO()
            try:
                pil_img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
            except Exception:  # noqa: BLE001
                continue
            new_blob = buf.getvalue()

            if len(new_blob) < len(blob):
                saved_bytes += len(blob) - len(new_blob)
                compressed_count += 1
                image_part._blob = new_blob

            # Büyük görsellerde piksel arabelleklerini hemen serbest
            # bırak; Python'ın çöp toplayıcısını açıkça tetiklemek,
            # özellikle çok yüksek çözünürlüklü görsellerin art arda
            # işlendiği durumlarda tepe bellek kullanımını düşürür.
            del pil_img, buf, blob, new_blob
            gc.collect()

    prs.save(output_path)
    return compressed_count, saved_bytes


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
        ext = os.path.splitext(input_path)[1].lower()
        is_pptx_like = ext in (".pptx", ".pptm", ".potx")

        raw_size_mb = os.path.getsize(input_path) / (1024 * 1024)
        already_shrunk = False

        # 0) Büyük dosya: python-pptx'e DOKUNMADAN önce zip seviyesinde
        #    küçült. python-pptx bir dosyayı açarken TAMAMINI belleğe alır;
        #    500MB'lık bir sunumu (hele her parça için yeniden) açmak
        #    sunucunun belleğini taşırıyordu. Akış hâlindeki bu ön adım
        #    aynı anda yalnızca tek bir görseli bellekte tutar ve sonraki
        #    bütün adımlar küçülmüş dosya üzerinde çalışır.
        if is_pptx_like and raw_size_mb > IMAGE_COMPRESS_THRESHOLD_MB:
            await status_msg.edit_text(
                f"🗜️ Dosya {raw_size_mb:.0f} MB, görseller tek tek "
                "küçültülüyor... (birkaç dakika sürebilir)"
            )
            passes = (
                (IMAGE_MAX_DIMENSION_PX, IMAGE_JPEG_QUALITY),
                (IMAGE_MAX_DIMENSION_PX_AGGRESSIVE, IMAGE_JPEG_QUALITY_AGGRESSIVE),
            )
            for pass_no, (max_dim, quality) in enumerate(passes, start=1):
                shrunk_path = os.path.join(work_dir, f"shrunk{pass_no}_{file_name}")
                try:
                    img_count, media_count = await loop.run_in_executor(
                        None,
                        shrink_pptx_streaming,
                        input_path,
                        shrunk_path,
                        max_dim,
                        quality,
                    )
                except Exception:  # noqa: BLE001
                    logger.exception(
                        "Akış hâlinde küçültme başarısız, mevcut dosyayla devam"
                    )
                    try:
                        os.remove(shrunk_path)
                    except OSError:
                        pass
                    break

                new_size_mb = os.path.getsize(shrunk_path) / (1024 * 1024)
                logger.info(
                    "Küçültme turu %d: %d görsel, %d video/ses, %.0fMB -> %.0fMB",
                    pass_no, img_count, media_count, raw_size_mb, new_size_mb,
                )
                # Eski (büyük) dosyayı hemen sil; disk de sınırlı.
                try:
                    os.remove(input_path)
                except OSError:
                    pass
                input_path = shrunk_path
                raw_size_mb = new_size_mb
                already_shrunk = True
                if new_size_mb <= IMAGE_SECOND_PASS_THRESHOLD_MB:
                    break

        # Parçalama gerekip gerekmeyeceğine ERKEN karar veriyoruz (tüm
        # dosyayı sıkıştırmadan/font taramadan ÖNCE). Çünkü büyük
        # dosyalarda bu "iyileştirme" adımlarını TÜM dosya üzerinde
        # çalıştırmak bellek taşmasına neden olabiliyor — bunun yerine
        # önce parçalara bölüp, her adımı (sıkıştırma dahil) yalnızca
        # o küçük parça üzerinde çalıştırıyoruz.
        slide_count = 0
        if is_pptx_like:
            try:
                slide_count = count_slides_fast(input_path)
            except Exception:  # noqa: BLE001
                logger.exception("Slayt sayısı okunamadı")

        # Fontlar: parçalansın ya da parçalanmasın, dosyanın TAMAMI için bir
        # kez hazırlanır. Tarama yalnızca zip içindeki XML'leri okur, büyük
        # dosyada da ucuzdur. Eksik font = farklı harf genişliği = metnin
        # ve çevresindeki görsellerin kayması.
        if is_pptx_like:
            try:
                await loop.run_in_executor(None, ensure_fonts_available, input_path)
            except Exception:  # noqa: BLE001
                logger.exception("Font hazırlığı sırasında hata oluştu, devam ediliyor")

        will_chunk = is_pptx_like and (
            slide_count > CHUNK_SLIDE_THRESHOLD
            or raw_size_mb > CHUNK_FILE_SIZE_MB_THRESHOLD
        )

        convert_input_path = input_path

        if will_chunk:
            # Büyük dosya: kalan görsel sıkıştırma işlemi her parça
            # üzerinde ayrı ayrı (çok daha az bellekle) uygulanacak.
            await status_msg.edit_text(
                f"🔄 Büyük dosya tespit edildi ({slide_count} slayt, "
                f"{raw_size_mb:.0f} MB).\nParçalara bölünüp, her parça "
                "ayrı ayrı sıkıştırılıp dönüştürülecek..."
            )
        else:
            # Küçük/orta boy dosya: gerekiyorsa görselleri sıkıştır.
            if raw_size_mb > IMAGE_COMPRESS_THRESHOLD_MB and is_pptx_like:
                await status_msg.edit_text(
                    f"🗜️ Dosya {raw_size_mb:.0f}MB, görseller sıkıştırılıyor..."
                )
                compressed_path = os.path.join(work_dir, "compressed_" + file_name)
                try:
                    compressed_count, saved_bytes = await loop.run_in_executor(
                        None, compress_pptx_images, convert_input_path, compressed_path
                    )
                    if compressed_count > 0 and os.path.exists(compressed_path):
                        convert_input_path = compressed_path
                        new_size_mb = os.path.getsize(convert_input_path) / (1024 * 1024)
                        logger.info(
                            "%d görsel sıkıştırıldı, %.0fMB -> %.0fMB",
                            compressed_count, raw_size_mb, new_size_mb,
                        )
                        if new_size_mb > IMAGE_SECOND_PASS_THRESHOLD_MB:
                            await status_msg.edit_text(
                                f"🗜️ Hâlâ {new_size_mb:.0f}MB, daha agresif "
                                "sıkıştırma ile ikinci tur deneniyor..."
                            )
                            compressed_path2 = os.path.join(
                                work_dir, "compressed2_" + file_name
                            )
                            try:
                                count2, saved2 = await loop.run_in_executor(
                                    None,
                                    compress_pptx_images,
                                    convert_input_path,
                                    compressed_path2,
                                    IMAGE_MAX_DIMENSION_PX_AGGRESSIVE,
                                    IMAGE_JPEG_QUALITY_AGGRESSIVE,
                                )
                                if count2 > 0 and os.path.exists(compressed_path2):
                                    convert_input_path = compressed_path2
                            except Exception:  # noqa: BLE001
                                logger.exception("İkinci sıkıştırma turu başarısız")
                except Exception:  # noqa: BLE001
                    logger.exception(
                        "Görsel sıkıştırma başarısız, orijinal dosyayla devam ediliyor"
                    )

        # Çok büyük dosyalarda LibreOffice'in her seferinde işlediği yükü
        # daha da azaltmak için daha küçük parçalar kullan.
        effective_chunk_size = (
            2 if raw_size_mb > MEMORY_SAFE_PPTX_MB_THRESHOLD else CHUNK_SIZE
        )
        use_chunking = will_chunk

        if use_chunking:
            progress_cb = _make_progress_callback(loop, status_msg)
            pdf_path = await loop.run_in_executor(
                None,
                convert_pptx_to_pdf_chunked,
                convert_input_path,
                work_dir,
                effective_chunk_size,
                progress_cb,
                already_shrunk,
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
        error_text = str(exc).lower()
        if (
            "413" in error_text
            or "request entity too large" in error_text
            or "payload too large" in error_text
        ):
            await status_msg.edit_text(
                "❌ Oluşan PDF, Telegram'ın gönderme sınırını (50MB) aşıyor "
                "ve bir sayfası tek başına bu sınırın üzerinde olduğu için "
                "otomatik bölme de yetersiz kaldı.\n\n"
                "Bu genelde tek bir slaytta aşırı yüksek çözünürlüklü bir "
                "görsel olduğunu gösterir. O slaytı ayrı gönderip dener "
                "misin, ya da kaynaktaki görseli küçültüp tekrar dener misin?"
            )
        else:
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
        if (
            "too big" in error_text
            or "file is too big" in error_text
            or "413" in error_text
            or "request entity too large" in error_text
            or "payload too large" in error_text
        ):
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
