"""Mail ile indirilen RPT Excel dosyalarını eski düzene uygun şekilde birleştirir."""
from __future__ import annotations

from pathlib import Path
import shutil
import pandas as pd

from config import PROJECT_ROOT
from app.helpers import today_str_en, ensure_folder


# Eski sistemde kullandığın düzenli sütun listesi
TARGET_COLUMNS = [
    "Fatura Gönderi Kodu",
    "Oluşturulma Tarihi",
    "Fatura Numarası",
    "Fatura Tarihi",
    "Toplam Fatura Tutarı",
    "Toplam Fatura Kdv",
    "Faturayı Düzenleyen Birim",
    "Müşteri Adı",
    "Gönderici Müşteri Kodu",
    "Gönderici Müşteri",
    "Gönderen Müşteri Adresi",
    "Gönderen Müşteri Telefon Numarası",
    "Alıcı Müşteri Kodu",
    "Alıcı Müşteri",
    "Alıcı Müşteri Adres",
    "Alıcı Müşteri Telefon Numarası",
    "Çıkış Birimi",
    "Çıkış İli",
    "Çıkış Tarihi",
    "Varış Birimi",
    "Varış İli",
    "Varış Tarihi",
    "Kargo Tipi",
    "Ödeme Tipi",
    "Alım Tipi",
    "Teslim Birimi",
    "Teslim İli",
    "Ürün Adı",
    "Toplam Kargo Adedi",
    "Desi / Kg",
    "Fatura Tipi",
    "Gönderi Kodu",
    "İrsaliye Numarası",
    "Ürün Bedeli",
    "İrsaliye Matrahı",
    "Kdv",
    "İrsaliye Matrahı+KDV",
    "Kargo Statüsü",
    "Kargo Statü Detayı",
    "Mesafe (Km)",
    "Mesafe Açıklaması",
    "Teslim Alan",
    "Teslim Tarihi",
    "Teslim Saati",
    "Ambar Tesellüm",
    "Sevk İrsaliye No.",
    "Bilgi",
    "Açıklama",
    "Tutanak Numarası",
    "Özel Alan",
    "Gönderici Segment Kodu",
    "Gönderici Segment Adı",
    "Alıcı Segment Kodu",
    "Alıcı Segment Adı",
    "Fatura Posta Hizmet Bedeli",
    "İrsaliye Posta Hizmet Bedeli",
    "YKPlus Mı?",
    "YKPlus Tipi",
]


def merge_rpt_excels() -> Path | None:
    """Downloads klasöründeki RPT Excel dosyalarını toplayıp,
    eski sistemdeki gibi temiz birleştirilmiş Excel üretir.
    """

    # 📂 Kullanıcının indirmeler klasörü
    downloads = Path.home() / "Downloads"

    # 📌 Bugünün klasörü V3 içinde
    today = today_str_en()
    target_folder = PROJECT_ROOT / "MailExcels" / today
    ensure_folder(target_folder)

    # 🔍 RPT formatındaki Excel dosyalarını bul
    rpt_files = list(downloads.glob("RPT*.xls")) + list(downloads.glob("RPT*.xlsx"))

    if not rpt_files:
        print("❌ Downloads klasöründe RPT Excel dosyası bulunamadı.")
        return None

    print(f"📥 Bulunan RPT dosyası sayısı: {len(rpt_files)}")

    merged_df_list: list[pd.DataFrame] = []

    for file in rpt_files:
        print(f"➡ Kopyalanıyor: {file.name}")

        # 📁 V3 içindeki günlük klasöre kopyala (arşiv için)
        dest = target_folder / file.name
        shutil.copy2(file, dest)

        # 📄 Eski sistemle uyumlu okuma:
        # - dtype=str → tip karışıklığı olmasın
        # - skiprows=4 → üstteki açıklama satırlarını at
        try:
            df_raw = pd.read_excel(dest, dtype=str, skiprows=4)

            # Sadece ihtiyacımız olan sütunları al
            df = df_raw[TARGET_COLUMNS].copy()

            # Tamamen boş satırları temizle
            df.dropna(how="all", inplace=True)

            merged_df_list.append(df)
            print(f"   ✔ Okundu ve temizlendi: {file.name}")
        except Exception as exc:
            print(f"   ❌ Okuma / temizleme hatası: {file.name} ({exc})")

    if not merged_df_list:
        print("❌ Hiçbir Excel düzgün okunamadı! Birleştirme yapılamıyor.")
        return None

    # 🧩 Birleştir (eski mantıkla)
    merged_df = pd.concat(merged_df_list, ignore_index=True)

    # 📁 Çıktı dosyası (eski dosya adınla aynı formatta)
    output_path = target_folder / f"{today} - Birleştirilmiş Mail Excelleri.xlsx"

    merged_df.to_excel(output_path, index=False)

    print(f"✅ Birleştirme tamamlandı → {output_path}")

    return output_path


def main() -> None:
    merge_rpt_excels()


if __name__ == "__main__":
    main()
