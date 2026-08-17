import pandas as pd
import json
import os

# Betiğin çalıştığı klasörün tam yolunu otomatik bulur (arayuz/mantik klasörü)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KLASOR = os.path.join(BASE_DIR, "katalog_veri_json")

# Dosya yolları artık tamamen bilgisayardan bağımsız, dinamik hale geldi
df = pd.read_excel(os.path.join(KLASOR, 'ozel_kw_verileri.xlsx'), header=0)
excel_yolu_flans = pd.read_excel(os.path.join(KLASOR, 'standart_flans_verileri.xlsx'), header=0)
mil_olculeri = pd.read_excel(os.path.join(KLASOR, 'kesilmis_mil_caplari.xlsx'), header=0)
govde_kod_yasaklari = pd.read_excel(os.path.join(KLASOR, 'govde_olusturulamayacak kodlar.xlsx'), header=0)

govde_kod_yasaklari.columns = govde_kod_yasaklari.columns.str.strip().str.upper()

katalog_dict = {}
flans_olculeri = {}
cap_dict = {}
govde_dict = {}

# --- 1. KATALOG VERİLERİ ---
for index, row in df.iterrows():
    ham_kod = str(row['kod'])
    ham_kw = str(row['kw'])
    ham_devir = str(row['devir'])

    if ham_kod.lower() == 'nan':
        continue

    sifre = ham_kod.replace(" ", "").strip()
    kw = ham_kw.strip()
    devir = ham_devir.strip()

    katalog_dict[sifre] = {"kw": kw, "devir": devir}

with open(os.path.join(KLASOR, 'katalog.json'), 'w', encoding='utf-8') as f:
    json.dump(katalog_dict, f, indent=4)

print(f"Boşluklar gitti! Toplam {len(katalog_dict)} kayıt boşluksuz şekilde katalog.json dosyasına yazıldı.")


# --- 2. STANDART FLANŞ VERİLERİ ---
for i, row_flans in excel_yolu_flans.iterrows():
    ham_M = str(row_flans['m'])
    ham_yapi_sekli = str(row_flans['yapi_sekli'])
    ham_flans_olcusu = str(row_flans['flans_olcusu'])
    ham_flans_kod = str(row_flans['kod'])
    ham_bilgi = str(row_flans['standart_bilgi'])

    if ham_flans_kod.lower() == 'nan':
        continue

    if ham_flans_olcusu.lower() == 'nan':
        ham_flans_olcusu = ham_flans_olcusu.replace("nan", "YOK").strip()

    flans_sifre = ham_flans_kod.replace(" ", "").strip()
    m = ham_M.strip()
    yapi_sekli = ham_yapi_sekli.strip()
    flans_olcusu = ham_flans_olcusu.strip()
    bilgi = ham_bilgi.strip()

    flans_olculeri[flans_sifre] = {"M": m, "Yapi Sekli": yapi_sekli, "flans olcusu": flans_olcusu, "standart bilgi": bilgi}

with open(os.path.join(KLASOR, 'standart_flans.json'), 'w', encoding='utf-8') as f:
    json.dump(flans_olculeri, f, indent=4)

print(f"Boşluklar gitti! Toplam {len(flans_olculeri)} kayıt boşluksuz şekilde standart_flans.json dosyasına yazıldı.")


# --- 3. MİL ÖLÇÜLERİ ---
mil_olculeri.columns = ["grup", "cap"]
for index, row in mil_olculeri.iterrows():
    grup = str(row['grup']).strip().upper()
    cap = str(row['cap']).strip().upper()

    if grup != 'NAN' and cap != 'NAN':
        cap_dict[grup] = cap

with open(os.path.join(KLASOR, 'cap.json'), 'w', encoding='utf-8') as f:
    json.dump(cap_dict, f, indent=4)

print(f"Başarılı! Toplam {len(cap_dict)} adet mil çapı cap.json dosyasına yazıldı.")


# --- 4. YASAKLI GÖVDE KODLARI ---
for index, row in govde_kod_yasaklari.iterrows():
    yasakli_kod = str(row['ACIKLAMA']).strip().upper()

    if yasakli_kod != 'NAN':
        govde_dict[yasakli_kod] = True

with open(os.path.join(KLASOR, 'govde_yasaklari.json'), 'w', encoding='utf-8') as f:
    json.dump(govde_dict, f, indent=4)

print(f"Yasaklı gövdeler eklendi! Toplam {len(govde_dict)} kayıt govde_yasaklari.json dosyasına yazıldı.")