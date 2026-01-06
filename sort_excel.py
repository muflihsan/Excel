import pandas as pd
import re

# ===============================
# LOAD FILE EXCEL
# ===============================
file_path = "farmasi.xlsx"   # ganti sesuai nama file
df = pd.read_excel(file_path)

# ===============================
# NORMALISASI KOLOM
# ===============================
df.columns = df.columns.str.strip().str.upper()

# BUANG KOLOM KOSONG / HANTU
df = df.dropna(axis=1, how="all")
df = df.loc[:, df.columns != ""]

total_rows = len(df)

# ===============================
# LIST 38 PROVINSI INDONESIA
# ===============================
PROVINCES = [
    "Aceh","Sumatera Utara","Sumatera Barat","Riau","Kepulauan Riau","Jambi",
    "Sumatera Selatan","Bangka Belitung","Bengkulu","Lampung","DKI Jakarta",
    "Jawa Barat","Jawa Tengah","DI Yogyakarta","Jawa Timur","Banten","Bali",
    "Nusa Tenggara Barat","Nusa Tenggara Timur",
    "Kalimantan Barat","Kalimantan Tengah","Kalimantan Selatan",
    "Kalimantan Timur","Kalimantan Utara",
    "Sulawesi Utara","Sulawesi Tengah","Sulawesi Selatan","Sulawesi Tenggara",
    "Gorontalo","Sulawesi Barat",
    "Maluku","Maluku Utara",
    "Papua","Papua Barat","Papua Barat Daya","Papua Tengah",
    "Papua Pegunungan","Papua Selatan"
]

# ===============================
# DEBUG COUNTER
# ===============================
city_found = 0
province_found = 0

# ===============================
# FUNCTION: EXTRACT KOTA / KABUPATEN (BERSIH)
# ===============================
def extract_city(address):
    global city_found
    if pd.isna(address):
        return ""

    text = str(address)

    patterns = [
        r"Kota\s+[A-Za-z\s]+",
        r"Kab\.\s*[A-Za-z\s]+",
        r"Kabupaten\s+[A-Za-z\s]+"
    ]

    for p in patterns:
        match = re.search(p, text)
        if match:
            city = match.group(0).strip()

            # Normalisasi Kab. -> Kabupaten
            city = re.sub(r"^Kab\.", "Kabupaten", city)

            # Hapus nama provinsi jika ikut terbaca
            for prov in PROVINCES:
                city = re.sub(rf"\s+{prov}$", "", city, flags=re.IGNORECASE)

            city_found += 1
            return city.strip()

    return ""

# ===============================
# FUNCTION: EXTRACT PROVINSI
# ===============================
def extract_province(address):
    global province_found
    if pd.isna(address):
        return ""

    text = str(address).lower()
    for prov in PROVINCES:
        if prov.lower() in text:
            province_found += 1
            return prov

    return ""

# ===============================
# PROCESS DATA WITH CLEAN PROGRESS
# ===============================
kota_list = []
prov_list = []

last_progress = -1

for idx, alamat in enumerate(df["ALAMAT"], start=1):
    kota_list.append(extract_city(alamat))
    prov_list.append(extract_province(alamat))

    progress = int((idx / total_rows) * 100)
    if progress % 10 == 0 and progress != last_progress:
        print(f"Progress: {progress}% ({idx}/{total_rows})")
        last_progress = progress

# ===============================
# ASSIGN HASIL
# ===============================
df["KOTA_KABUPATEN"] = kota_list
df["PROVINSI"] = prov_list

# ===============================
# SAVE OUTPUT
# ===============================
output_file = "farmasi_hasil.xlsx"
df.to_excel(output_file, index=False)

# ===============================
# FINAL DEBUG SUMMARY
# ===============================
print("\n=== DEBUG SUMMARY ===")
print(f"Total data               : {total_rows}")
print(f"Kota/Kabupaten terdeteksi: {city_found} ({city_found/total_rows*100:.2f}%)")
print(f"Provinsi terdeteksi      : {province_found} ({province_found/total_rows*100:.2f}%)")
print(f"File output              : {output_file}")
print("Selesai ✔")
