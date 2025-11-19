"""
Yurtiçi Kargo PDF e-faturalarındaki QR kodlarını okuyup Excel çıktısı üreten modül.
"""

from __future__ import annotations

import io
import json
from pathlib import Path
from typing import List, Iterable

import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
from pyzbar.pyzbar import decode

from config import Z_EFATURA_FOLDER, YURTICI_PREFIXES, PROJECT_ROOT
from app.helpers import ensure_folder, today_str_en


def find_qr_in_pdf(pdf_path: Path) -> List[dict]:
    """Tek bir PDF içindeki QR kodları çözüp JSON liste halinde döndürür."""
    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"❌ PDF açılamadı: {pdf_path} ({exc})")
        return []

    results = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        images = page.get_images(full=True)

        for xref, *_ in images:
            base = doc.extract_image(xref)
            img_bytes = base["image"]

            image = Image.open(io.BytesIO(img_bytes))
            decoded = decode(image)

            for qr in decoded:
                raw = qr.data.decode("utf-8")

                try:
                    qr_json = json.loads(raw)
                    qr_json["Dosya Adı"] = pdf_path.name
                    results.append(qr_json)
                except json.JSONDecodeError:
                    print(f"⚠️ JSON olmayan QR bulundu: {raw}")
                    results.append({"Dosya Adı": pdf_path.name, "QR_Data": raw})

    return results


def process_all_pdfs() -> Path | None:
    """PDF klasöründeki tüm YKA/YKB dosyalarını tarar ve Excel çıktısı üretir."""

    if not Z_EFATURA_FOLDER.exists():
        print(f"❌ Klasör yok: {Z_EFATURA_FOLDER}")
        return None

    pdf_files = [
        f for f in Z_EFATURA_FOLDER.iterdir()
        if f.suffix.lower() == ".pdf" and f.name.startswith(tuple(YURTICI_PREFIXES))
    ]

    if not pdf_files:
        print("🔍 Uygun PDF bulunamadı (YKA20 / YKB20).")
        return None

    all_data = []

    print("📌 Toplam PDF:", len(pdf_files))

    for pdf in pdf_files:
        print(f"📄 İşleniyor: {pdf.name}")
        qr_list = find_qr_in_pdf(pdf)
        all_data.extend(qr_list)

    if not all_data:
        print("⚠️ QR bulunamadı!")
        return None

    df = pd.DataFrame(all_data)

    # Aynı QR "no" bir daha gelirse tekrarı sil
    if "no" in df.columns:
        df.drop_duplicates(subset=["no"], inplace=True)

    output = PROJECT_ROOT / "Duzenlenen_QR_listesi.xlsx"
    df.to_excel(output, index=False)

    print(f"✅ QR kodları kaydedildi → {output}")
    return output


def main():
    process_all_pdfs()


if __name__ == "__main__":
    main()
