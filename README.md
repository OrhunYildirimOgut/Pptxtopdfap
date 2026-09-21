# PPTX → PDF Telegram Botu

Gönderilen `.pptx` dosyalarını, görsel ve metin kalitesi bozulmadan
PDF'e dönüştüren bir Telegram botu.

## Neden LibreOffice?

Dönüştürme `soffice --headless` (LibreOffice) ile yapılır. LibreOffice
sunumu gerçek bir ofis motoruyla render ettiği için:

- Görseller yeniden sıkıştırılmaz / küçültülmez (`ReduceImageResolution=false`, `Quality=100`)
- Metinler vektörel kalır (bulanıklaşmaz, piksel piksel görünmez)
- Fontlar, animasyon dışı düzen ve renkler orijinaline en yakın şekilde korunur

`python-pptx` ile manuel PDF üretmek yerine bu yöntem tercih edildi,
çünkü python-pptx bir render motoru değildir ve kaliteli/doğru bir PDF
çıktısı üretmez.

## Kurulum

```bash
# 1) Sistem bağımlılığı: LibreOffice
sudo apt-get update
sudo apt-get install -y libreoffice

# 2) Python bağımlılıkları
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Botu oluşturma

1. Telegram'da [@BotFather](https://t.me/BotFather) ile konuşup `/newbot` komutunu çalıştır.
2. Sana verilen token'ı kopyala.

## Çalıştırma

```bash
export TELEGRAM_BOT_TOKEN="123456:ABC-DEF..."
python3 bot.py
```

Bot polling modunda çalışır; terminali açık tutman (veya bir sunucuda
`systemd` / `screen` / `tmux` / `pm2` ile arka planda çalıştırman) yeterli.

## Kullanım

1. Telegram'da botunu bul, `/start` yaz.
2. Bir `.pptx` dosyasını **dosya (document)** olarak gönder.
3. Bot dosyayı indirir, LibreOffice ile PDF'e çevirir ve sana geri yollar.

## Notlar / Sınırlamalar

- PowerPoint'in metni kutuya sığdırma ayarları (`fontScale` ve
  `lnSpcReduction`) korunur ve LibreOffice tarafından uygulanır. Şablondan
  gelen punto değerlerini elle değiştirmek veya otomatik sığdırmayı
  kapatmak, metnin büyüyüp görsellerin altına taşmasına neden olabilir.
  Eksik fontların ikamesi yine de görünüm farkı oluşturabilir; her sunum
  için PowerPoint ile piksel düzeyinde aynı sonuç garanti edilmez.
- Varsayılan Telegram Bot API, botların indirebileceği dosya boyutunu
  sınırlar (genelde 20 MB civarı). Daha büyük dosyalar için kendi
  [Local Bot API Server](https://github.com/tdlib/telegram-bot-api)
  kurulumunu kullanman gerekir.
- Animasyonlar ve geçiş efektleri PDF formatında doğası gereği
  desteklenmez (PDF statik bir sayfa formatıdır); slaytların son
  görünümü (görsel + metin) korunur.
- Eşzamanlı (aynı anda birden fazla kullanıcıdan gelen) istekler için
  her dönüştürme ayrı bir geçici LibreOffice profiliyle çalışır, bu
  yüzden çakışma yaşanmaz.
- Sunucuda çalıştırırken bot'u arka planda tutmak için örnek bir
  systemd servis dosyası isteyebilirsin.

## Test

```bash
python -m unittest discover -s tests -v
```

Testler Telegram'a bağlanmaz ve LibreOffice gerektirmez. Küçük dosya,
slaytlara bölme ve görsel sıkıştırma yollarında metin ayarlarının
dönüştürücüye değişmeden ulaştığını kontrol eder. Görsel doğrulama için
aynı sunumu Docker imajındaki LibreOffice ile PDF'e dönüştürüp karşılaştırın.
