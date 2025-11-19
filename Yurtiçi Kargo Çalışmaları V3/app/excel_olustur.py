"""QR sonuçlarından Yurtiçi Kargo'ya gönderilecek sade Excel'i oluşturur."""
from __future__ import annotations

from pathlib import Path
import pandas as pd

from config import PROJECT_ROOT, DEFAULT_EMAIL
from app.helpers import today_str_en


def create_yurtici_excel() -> Path | None:
    """Duzenlenen_QR_listesi.xlsx dosyasından günlük gönderim Excel'ini üretir."""
    # V3 klasöründeki QR sonuç dosyamız
    qr_excel_path = PROJECT_ROOT / "Duzenlenen_QR_listesi.xlsx"

    if not qr_excel_path.exists():
        print(f"❌ Hata: QR sonuç Excel'i bulunamadı! ({qr_excel_path})")
        return None

    try:
        df = pd.read_excel(qr_excel_path)
    except Exception as exc:
        print(f"❌ Hata: QR sonuç Excel'i okunamadı! ({exc})")
        return None

    if df.empty:
        print("⚠️ Uyarı: QR sonuç dosyası boş, oluşturulacak satır yok.")
        return None

    # ▶ V2 otomasyonunun beklediği 'mail' sütununu ekliyoruz
    df["mail"] = DEFAULT_EMAIL  # örn: yhacioglu@interkan.com.tr

    # 📌 ÇIKTI DAİMA MASAÜSTÜNE KAYDEDİLECEK
    today = today_str_en()
    output_path = Path.home() / "Desktop" / f"{today} - Yurtiçi Kargo Faturaları.xlsx"

    try:
        df.to_excel(output_path, index=False)
    except Exception as exc:
        print(f"❌ Hata: Gönderim Excel'i yazılamadı! ({exc})")
        return None

    print(f"✅ Yeni Excel dosyası oluşturuldu → {output_path}")
    return output_path


def main() -> None:
    create_yurtici_excel()


if __name__ == "__main__":
    main()
