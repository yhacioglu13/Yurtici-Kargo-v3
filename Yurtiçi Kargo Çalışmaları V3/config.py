"""V3 otomasyon projesi için tüm sabit yollar ve ayarlar."""

from pathlib import Path

# ---------------------------------------------------------
# 1) Projenin kök klasörü
# ---------------------------------------------------------
PROJECT_ROOT = Path(
    r"C:\Users\SL02.POLINT\Desktop\MASAÜSTÜ\Yazılım Çalışmalarım\Yurtiçi Kargo Çalışmaları V3"
)

# ---------------------------------------------------------
# 2) PDF e-faturaların bulunduğu Z: klasörü
# ---------------------------------------------------------
Z_EFATURA_FOLDER = Path(
    r"Z:\SATINALMA - LOJİSTİK\ÖDEMELER\KONTROL EDİLECEK E-FATURALAR\İNTERKAN"
)

# ---------------------------------------------------------
# 3) Masaüstü klasörü
# ---------------------------------------------------------
DESKTOP = Path(r"C:\Users\SL02.POLINT\Desktop")

# ---------------------------------------------------------
# 4) QR çıktısı ve gönderilecek Excel için yollar
# ---------------------------------------------------------
QR_OUTPUT_EXCEL = PROJECT_ROOT / "Duzenlenen_QR_listesi.xlsx"

# Masaüstüne oluşturulacak Yurtiçi Excel formatı:
#   örnek: "19.November.2025 - Yurtiçi Kargo Faturaları.xlsx"
DESKTOP_OUTPUT_EXCEL_TEMPLATE = "{today} - Yurtiçi Kargo Faturaları.xlsx"

# ---------------------------------------------------------
# 5) Mail RPT excellerinin işleneceği klasör
# ---------------------------------------------------------
MAIL_EXCEL_ROOT = PROJECT_ROOT / "MailExcels"

# ---------------------------------------------------------
# 6) Tarife (fiyat) Excel dosyası
#    → Bu dosyanın şu klasörde bulunması gerekiyor:
#      V3\Tarife\Yurt İçi Kargo Fiyatları.xlsx
# ---------------------------------------------------------
TARIFF_EXCEL_PATH = PROJECT_ROOT / "Tarife" / "Yurt İçi Kargo Fiyatları.xlsx"

# ---------------------------------------------------------
# 7) Sonuç excelleri (maliyet kontrolü sonrası fatura bazlı)
# ---------------------------------------------------------
RESULT_EXCEL_ROOT = PROJECT_ROOT / "Results"

# ---------------------------------------------------------
# 8) Yurtiçi PDF prefixleri
# ---------------------------------------------------------
YURTICI_PREFIXES = ("YKA20", "YKB20")

# ---------------------------------------------------------
# 9) ChromeDriver yolu (otomasyon için)
# ---------------------------------------------------------
CHROMEDRIVER_PATH = Path(
    r"C:\Users\SL02.POLINT\Desktop\MASAÜSTÜ\Yazılım Çalışmalarım\Yurtiçi Kargo Çalışmaları V2\google güncel chromedriver\chromedriver-win64\chromedriver.exe"
)

# ---------------------------------------------------------
# 10) Varsayılan mail adresi
# ---------------------------------------------------------
DEFAULT_EMAIL = "yhacioglu@interkan.com.tr"

