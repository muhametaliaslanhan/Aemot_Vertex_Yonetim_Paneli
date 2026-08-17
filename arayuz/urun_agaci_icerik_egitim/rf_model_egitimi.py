import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split
import joblib
import os
from sklearn.metrics import accuracy_score, f1_score, precision_score, recall_score

dosya_yolu = r"C:\Users\aemot.stajyer\Desktop\urun_agaci_verileri\urun_agaci_verileri_bir.xlsx"
kayit_klasoru = r"C:\Users\aemot.stajyer\Desktop\AEMOT_VERTEX_YonetimPaneli\arayuz\urun_agaci_icerik_egitim"

if not os.path.exists(kayit_klasoru):
    os.makedirs(kayit_klasoru, exist_ok=True)

print("1/8: Veri seti okunuyor...")
df = pd.read_excel(dosya_yolu)

print("2/8: Boş hücreler temizleniyor ve tipler düzenleniyor...")
df = df.dropna(subset=['Tepekodu1', 'Bılesen Malzeme Kodu', 'Bılesen Malzeme Acıklama', 'Level'])
df['Tepekodu1'] = df['Tepekodu1'].astype(str)
df['Level'] = df['Level'].astype(int)
df['Montaj Basına'] = df['Montaj Basına'].fillna(0).astype(float)
if 'Sıra No' not in df.columns:
    df['Sıra No'] = 0
df['Sıra No'] = df['Sıra No'].fillna(0).astype(int)
if 'Çalışma Stn' not in df.columns:
    df['Çalışma Stn'] = ""
df['Çalışma Stn'] = df['Çalışma Stn'].fillna("").astype(str)

if 'Tepekodu' not in df.columns:
    df['Tepekodu'] = df['Tepekodu1']
if 'Tepekodu Acıklama' not in df.columns:
    df['Tepekodu Acıklama'] = ""

print("3/8: Tepekodu ve Açıklama özelliklere ayrıştırılıyor (Tüm detaylar)...")
df['Seri'] = df['Tepekodu1'].apply(lambda x: str(x)[0:4] if len(str(x)) > 4 else "UNK")
df['Kasa_Tipi'] = df['Tepekodu1'].apply(lambda x: str(x)[5:8] if len(str(x)) > 16 else "UNK")
df['Govde_Boyu'] = df['Tepekodu1'].apply(lambda x: str(x)[8] if len(str(x)) > 16 else "U")
df['Kutup'] = df['Tepekodu1'].apply(lambda x: str(x)[9:12] if len(str(x)) > 16 else "UNK")
df['Paket'] = df['Tepekodu1'].apply(lambda x: str(x)[12] if len(str(x)) > 16 else "U")
df['Yapi_Sekli'] = df['Tepekodu1'].apply(lambda x: str(x)[13] if len(str(x)) > 16 else "U")
df['Flans'] = df['Tepekodu1'].apply(lambda x: str(x)[14:16] if len(str(x)) > 16 else "UN")
df['Voltaj_Frekans'] = df['Tepekodu1'].apply(lambda x: str(x)[16:18] if len(str(x)) > 18 else "UN")
df['Opsiyon_1'] = df['Tepekodu1'].apply(lambda x: str(x)[18:22] if len(str(x)) > 21 else "UNK")
df['Opsiyon_2'] = df['Tepekodu1'].apply(lambda x: str(x)[22:] if len(str(x)) > 22 else "UNK")

# AÇIKLAMA EKLENDİ
df['Aciklama'] = df['Tepekodu Acıklama'].astype(str)

print("4/8: Motorlara ait ürün ağacı (BOM) profilleri oluşturuluyor...")
tepekodu_profilleri = {}
for tepekodu, group in df.groupby('Tepekodu1'):
    ilksatir = group.iloc[0]
    ozellikler = (ilksatir['Seri'], ilksatir['Kasa_Tipi'], ilksatir['Govde_Boyu'],
                  ilksatir['Kutup'], ilksatir['Paket'], ilksatir['Yapi_Sekli'],
                  ilksatir['Flans'], ilksatir['Voltaj_Frekans'], ilksatir['Opsiyon_1'], ilksatir['Opsiyon_2'], ilksatir['Aciklama'])

    bom_icerik = tuple(
        (row['Level'], row['Bılesen Malzeme Kodu'], row['Bılesen Malzeme Acıklama'],
         row['Montaj Basına'], row['Sıra No'], row['Çalışma Stn'], row['Tepekodu'], row['Tepekodu Acıklama'])
        for _, row in group.iterrows()
    )
    tepekodu_profilleri[tepekodu] = {'ozellikler': ozellikler, 'bom_icerik': bom_icerik, 'orijinal_df': group.copy()}

print("5/8: Benzersiz BOM profillerine Profil_ID atanıyor...")
benzersiz_bomlar = {}
profil_id_sayaci = 0
profil_mapping = {}
X_list = []
y_list = []

for tepekodu, data in tepekodu_profilleri.items():
    bom = data['bom_icerik']
    ozellikler = data['ozellikler']
    if bom not in benzersiz_bomlar:
        benzersiz_bomlar[bom] = profil_id_sayaci
        profil_mapping[profil_id_sayaci] = data['orijinal_df']
        profil_id_sayaci += 1
    X_list.append(ozellikler)
    y_list.append(benzersiz_bomlar[bom])

# KATEGORİK KOLONLARA 'Aciklama' EKLENDİ
kategorik_kolonlar = ['Seri', 'Kasa_Tipi', 'Govde_Boyu', 'Kutup', 'Paket', 'Yapi_Sekli', 'Flans', 'Voltaj_Frekans',
                      'Opsiyon_1', 'Opsiyon_2', 'Aciklama']
X_df = pd.DataFrame(X_list, columns=kategorik_kolonlar)
y_serisi = pd.Series(y_list)

print("6/8: X Değerleri LabelEncoder ile sayısal formata çevriliyor...")
encoderlar = {}
for kolon in kategorik_kolonlar:
    le = LabelEncoder()
    # Handle unseen labels by fitting with all unique values plus 'UNK' and 'U'
    all_classes = list(set(X_df[kolon].astype(str).unique().tolist() + ['UNK', 'U']))
    le.fit(all_classes)
    X_df[f'{kolon}_Kod'] = le.transform(X_df[kolon].astype(str))
    encoderlar[kolon] = le

X = X_df[[f"{k}_Kod" for k in kategorik_kolonlar]].astype('int16')
y = y_serisi

print("7/8: Model tüm veri setiyle eğitiliyor (Test ayırması iptal edildi)...")
rf_model = RandomForestClassifier(n_estimators=100, max_depth=None, random_state=42)
rf_model.fit(X, y)

# Tahminleri al
y_pred = rf_model.predict(X)

# Metrikleri hesapla (Weighted average kullanarak dengesiz dağılımı hesaba katıyoruz)
dogruluk = accuracy_score(y, y_pred)
f1 = f1_score(y, y_pred, average='weighted', zero_division=0)
kesinlik = precision_score(y, y_pred, average='weighted', zero_division=0)
duyarlilik = recall_score(y, y_pred, average='weighted', zero_division=0)

print("\n" + "="*50)
print(f"Doğruluk (Accuracy)  : % {dogruluk * 100:.2f}")
print(f"F1-Score (Weighted)  : % {f1 * 100:.2f}")
print(f"Kesinlik (Precision) : % {kesinlik * 100:.2f}")
print(f"Duyarlılık (Recall)  : % {duyarlilik * 100:.2f}")
print("="*50 + "\n")

print("8/8: Model, Encoderlar ve Profil Sözlüğü Kaydediliyor...")
joblib.dump(rf_model, os.path.join(kayit_klasoru, 'rf_model.pkl'))
joblib.dump(profil_mapping, os.path.join(kayit_klasoru, 'rf_profil_mapping.pkl'))

for kolon, le in encoderlar.items():
    joblib.dump(le, os.path.join(kayit_klasoru, f"le_{kolon.lower()}.pkl"))

print(f"İŞLEM TAMAM! Tüm dosyalar hazır:\n{kayit_klasoru}")