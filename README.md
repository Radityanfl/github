tolong buatkan saya file excel dari file json dan rapikan dengan menambahkan margin setiap kolom
Saya sudah mengubah file JSON Anda menjadi file Excel dan merapikannya dengan menambahkan margin (align center dan wrap text) di setiap kolom.

Anda bisa mengunduh hasilnya di sini:
📂 Raditya_Naufal_Wijaya_V3925033.xlsx 

Anda bilang:
tolong buatkan saya file excel dari file json
ChatGPT bilang:

Saya sudah mengubah file JSON Anda menjadi file Excel dengan tiap bagian berada di sheet terpisah.

Silakan unduh hasilnya di sini:
📂 Raditya_Naufal_Wijaya_V3925033.xlsx 
rapikan dengan menambah margin setiap kolom
ChatGPT bilang:

File Excel sudah dirapikan dengan menambah margin setiap kolom (kolom otomatis diperlebar, teks rata tengah, dan wrap text aktif).

Silakan unduh hasilnya di sini:
📂 Raditya_Naufal_Wijaya_V3925033_Rapih.xlsx
buatkan saya kode vscode yang mengubah file json ke excel 
ChatGPT bilang:

Baik, saya buatkan contoh kode Python yang bisa langsung Anda jalankan di VS Code untuk mengubah file JSON menjadi Excel dengan format rapi (auto width, margin kolom, rata tengah, dan wrap text).

Simpan file ini dengan nama misalnya json_to_excel.py:

import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

# Ganti path sesuai lokasi file JSON Anda
json_file = "Raditya_Naufal_Wijaya_V3925033.json"
output_file = "Raditya_Naufal_Wijaya_V3925033.xlsx"

def json_to_excel(json_path, excel_path):
    # Baca file JSON
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Buat workbook baru
    wb = Workbook()
    wb.remove(wb.active)  # hapus sheet default

    # Loop setiap bagian JSON untuk dibuatkan sheet
    for section, records in data.items():
        ws = wb.create_sheet(title=section[:30])  # judul max 31 karakter
        df = pd.DataFrame(records)

        # Tambahkan data frame ke worksheet
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Atur lebar kolom & alignment
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center",
                                           wrap_text=True)
            # tambah margin
            ws.column_dimensions[col_letter].width = max_length + 5

    # Simpan Excel
    wb.save(excel_path)
    print(f"File berhasil dibuat: {excel_path}")

if __name__ == "__main__":
    json_to_excel(json_file, output_file)

Cara menjalankan:

Simpan kode di file json_to_excel.py.

Letakkan file JSON di folder yang sama.

Jalankan di terminal VS Code:

python json_to_excel.py


Hasilnya berupa file Excel Raditya_Naufal_Wijaya_V3925033.xlsx.

Mau saya tambahkan fitur agar file JSON bisa dipilih lewat dialog file (GUI), jadi tidak perlu edit nama file manual?

Anda bilang:
import json
import pandas as pd

# Baca file JSON
with open("Raditya_Naufal_Wijaya_V3925033.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Ubah setiap bagian JSON menjadi CSV
for key, value in data.items():
    df = pd.DataFrame(value)
    # Buat nama file CSV dari nama key
    filename = key.replace(" ", "_")[:50] + ".csv"  # batasi panjang nama file
    df.to_csv(filename, index=False, encoding="utf-8-sig")
    print(f"File CSV berhasil dibuat: {filename}")
ChatGPT bilang:

Kode yang Anda buat sudah benar, hanya saja outputnya masih berupa CSV, bukan Excel (.xlsx).
Kalau ingin langsung menghasilkan file Excel dengan banyak sheet (seperti di permintaan sebelumnya), Anda bisa modifikasi kodenya seperti ini:

import json
import pandas as pd

# Baca file JSON
with open("Raditya_Naufal_Wijaya_V3925033.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Buat file Excel writer
with pd.ExcelWriter("Raditya_Naufal_Wijaya_V3925033.xlsx", engine="openpyxl") as writer:
    for key, value in data.items():
        df = pd.DataFrame(value)
        # Setiap key akan menjadi sheet di Excel
        sheet_name = key[:31]  # batasi panjang nama sheet max 31
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("File Excel berhasil dibuat: Raditya_Naufal_Wijaya_V3925033.xlsx")


👉 Bedanya:

pd.ExcelWriter digunakan untuk menggabungkan banyak DataFrame ke satu file Excel.

Setiap key di JSON jadi sheet terpisah di Excel.

Sheet name dibatasi maksimal 31 karakter sesuai aturan Excel.

Apakah Anda mau saya tambahkan juga pengaturan rapi (misalnya auto-width kolom, rata tengah, wrap text) langsung di kode ini biar hasil Excel lebih enak dibaca?
