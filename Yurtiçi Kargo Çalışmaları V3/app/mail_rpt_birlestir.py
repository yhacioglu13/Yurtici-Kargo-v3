"""Mail ile indirilen RPT Excel dosyalarını eski düzene uygun şekilde birleştirir."""
from __future__ import annotations

from pathlib import Path
import shutil
from shutil import SameFileError
import zipfile
import tkinter as tk
from tkinter import filedialog

import pandas as pd

from config import PROJECT_ROOT, DOWNLOADS_FOLDER
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


def ask_source_folder(default_folder: Path) -> Path | None:
    """Kullanıcıdan RPT dosyalarının olduğu klasörü seçmesini ister."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected = filedialog.askdirectory(
        title="RPT Excel/ZIP dosyalarının olduğu klasörü seçin",
        initialdir=str(default_folder),
    )
    root.destroy()

    if not selected:
        print("⚠️ Klasör seçimi iptal edildi.")
        return None

    chosen = Path(selected)
    print(f"📂 Seçilen klasör: {chosen}")
    return chosen


def _collect_rpt_excel_files(source_folder: Path, temp_extract_root: Path) -> list[Path]:
    """Klasördeki RPT Excel dosyalarını ve ZIP içindeki RPT Excel dosyalarını toplar."""
    excel_files = list(source_folder.glob("RPT*.xls")) + list(source_folder.glob("RPT*.xlsx"))

    zip_files = list(source_folder.glob("*.zip"))
    if zip_files:
        ensure_folder(temp_extract_root)

    for zip_path in zip_files:
        try:
            with zipfile.ZipFile(zip_path) as archive:
                members = [
                    member
                    for member in archive.namelist()
                    if Path(member).name.startswith("RPT")
                    and Path(member).suffix.lower() in {".xls", ".xlsx"}
                ]

                if not members:
                    continue

                for member in members:
                    extracted_path = Path(archive.extract(member, path=temp_extract_root))
                    excel_files.append(extracted_path)

                print(f"📦 ZIP içinden {len(members)} RPT dosyası çıkarıldı: {zip_path.name}")
        except Exception as exc:
            print(f"⚠️ ZIP okunamadı, atlandı: {zip_path.name} ({exc})")

    return excel_files


def merge_rpt_excels(source_folder: Path | None = None) -> Path | None:
    """Seçilen klasördeki RPT Excel dosyalarını toplayıp,
    eski sistemdeki gibi temiz birleştirilmiş Excel üretir.
    """

    # 📂 Varsayılan klasör config'den gelir
    source = source_folder or DOWNLOADS_FOLDER

    if not source.exists():
        print(f"❌ Kaynak klasör bulunamadı: {source}")
        return None

    # 📌 Bugünün klasörü V3 içinde
    today = today_str_en()
    target_folder = PROJECT_ROOT / "MailExcels" / today
    ensure_folder(target_folder)

    temp_extract_root = target_folder / "_tmp_zip_extract"
    rpt_files = _collect_rpt_excel_files(source, temp_extract_root)

    if not rpt_files:
        if temp_extract_root.exists():
            shutil.rmtree(temp_extract_root, ignore_errors=True)
        print(f"❌ {source} klasöründe RPT Excel veya ZIP içinde RPT dosyası bulunamadı.")
        return None

    print(f"📥 Bulunan toplam RPT dosyası sayısı: {len(rpt_files)}")

    merged_df_list: list[pd.DataFrame] = []

    for file in rpt_files:
        print(f"➡ İşleniyor: {file.name}")

        dest = target_folder / file.name
        try:
            shutil.copy2(file, dest)
        except SameFileError:
            pass

        try:
            df_raw = pd.read_excel(dest, dtype=str, skiprows=4)
            df = df_raw[TARGET_COLUMNS].copy()
            df.dropna(how="all", inplace=True)

            merged_df_list.append(df)
            print(f"   ✔ Okundu ve temizlendi: {file.name}")
        except Exception as exc:
            print(f"   ❌ Okuma / temizleme hatası: {file.name} ({exc})")

    if temp_extract_root.exists():
        shutil.rmtree(temp_extract_root, ignore_errors=True)

    if not merged_df_list:
        print("❌ Hiçbir Excel düzgün okunamadı! Birleştirme yapılamıyor.")
        return None

    merged_df = pd.concat(merged_df_list, ignore_index=True)
    output_path = target_folder / f"{today} - Birleştirilmiş Mail Excelleri.xlsx"
    merged_df.to_excel(output_path, index=False)

    print(f"✅ Birleştirme tamamlandı → {output_path}")
    return output_path


def main() -> None:
    merge_rpt_excels()


if __name__ == "__main__":
    main()
