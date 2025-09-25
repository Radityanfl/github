import pandas as pd
import json

# Ganti dengan nama file JSON kamu
input_file = "data.json"
output_file = "data.xlsx"

# Baca JSON
with open(input_file, "r", encoding="utf-8") as f:
    data = json.load(f)

# Jika JSON berbentuk array of objects (list of dicts)
if isinstance(data, list):
    df = pd.DataFrame(data)
else:
    # Kalau JSON nested, coba convert ke dataframe dengan normalisasi
    df = pd.json_normalize(data)

# Simpan ke Excel
df.to_excel(output_file, index=False, engine="openpyxl")

print(f"✅ File berhasil dikonversi ke {output_file}")