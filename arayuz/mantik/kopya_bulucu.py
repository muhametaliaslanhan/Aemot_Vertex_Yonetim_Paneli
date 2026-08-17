import pandas as pd
import collections

# Excel'i oku (kendi dosya yolunu kontrol et)
df = pd.read_excel(r'C:\Users\aemot.stajyer\Desktop\Muhammetali\aemot_masaustu\arayuz\mantik\katalog_veri_json\ozel_kw_verileri.xlsx', header=0)

sifreler = []

# Satır satır gez ve şifreleri listeye at
for index, row in df.iterrows():
    ham_kod = str(row['kod'])
    if ham_kod.lower() != 'nan':
        sifre = ham_kod.replace(" ", "").strip()
        sifreler.append(sifre)

# Hangi şifreden kaç tane var sayalım
sayac = collections.Counter(sifreler)

print("--- EXCELDE TEKRAR EDEN HATALI ŞİFRELER ---")
for sifre, adet in sayac.items():
    if adet > 1:
        print(f"Şifre: {sifre} - Excel'de {adet} defa yazılmış!")


def ayni_motor_var_mi(df, ozellikler):
    """Eklenmek istenen motor Excel'de BİREBİR aynı özelliklerle var mı kontrol eder.
    Varsa (True, motor_verisi) döner, yoksa (False, None) döner."""
    try:
        # Boş olan (NaN) hücreleri temizle ve string'e çevir
        for col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

        maske = (
                (df['VERİMLİLİK'] == ozellikler.get('VERİMLİLİK', '')) &
                (df['GRUP'] == ozellikler.get('GRUP', '')) &
                (df['GÖVDE BOYU'] == ozellikler.get('GÖVDE BOYU', '')) &
                (df['KUTUP'] == ozellikler.get('KUTUP', '')) &
                (df['PAKET BOYU'] == ozellikler.get('PAKET BOYU', '')) &
                (df['YAPI ŞEKLİ'] == ozellikler.get('YAPI ŞEKLİ', '')) &
                (df['FLANŞ ÖLÇÜSÜ'] == ozellikler.get('FLANŞ ÖLÇÜSÜ', '')) &
                (df['GÖVDE MALZEMESİ'] == ozellikler.get('GÖVDE MALZEMESİ', '')) &
                (df['EK AÇIKLAMA'] == ozellikler.get('EK AÇIKLAMA', ''))
        )

        if maske.any():
            # Eşleşen satırı bir sözlük (dictionary) olarak döndürüyoruz
            eslesen_satir = df[maske].iloc[0].to_dict()
            return True, eslesen_satir

        return False, None
    except Exception as e:
        print(f"Kopya bulucu hatası: {e}")
        return False, None