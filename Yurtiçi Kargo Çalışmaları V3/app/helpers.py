from pathlib import Path
from datetime import datetime

# Bugünün İngilizce formatlı tarih karşılığı → klasör/Excel adlarında kullanıyoruz.
def today_str_en() -> str:
    now = datetime.now()
    return now.strftime("%d.%B.%Y")  # örnek: "18.November.2025"

# Klasör yoksa oluşturan yardımcı fonksiyon
def ensure_folder(folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    return folder

# Güvenli şekilde Excel kolon temizliği (boşlukları kes, büyük/küçük harf eşitle)
def clean_col(col: str) -> str:
    return col.strip().replace("\n", " ").replace("\r", "")
