import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import accuracy_score
import joblib
import os
import json
import traceback

def excel_yolunu_bul():
    ana_dizin = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
    json_yolu = os.path.join(ana_dizin, "ayarlar.json")
    if os.path.exists(json_yolu):
        try:
            with open(json_yolu, "r", encoding="utf-8") as f:
                ayarlar = json.load(f)
                return ayarlar.get("excel_yolu", "")
        except Exception:
            pass
    return r"C:\Users\aemot.stajyer\Desktop\Muhammetali\KODOLUŞTURUCU2018GENELSOHEIL.xlsx"

def yapay_zeka_egit_rotor_saci():
    print("Rotor SACI eğitimi başlatılıyor...")

    excel_yolu = excel_yolunu_bul()

    if not os.path.exists(excel_yolu):
        print("HATA: Excel dosyası bulunamadı!")
        return

    try:
        excel_dosyasi = pd.ExcelFile(excel_yolu)
        hedef_sayfa = None

        # Sadece "ROTOR" yazarsak "ROTORLU MIL" sayfasını da yanlışlıkla çekebilir.
        # "ROTOR" ve "PAKET" kelimelerinin ikisini de aramak daha güvenli.
        for sayfa in excel_dosyasi.sheet_names:
            if "ROTOR" in sayfa.upper() and "SACI" in sayfa.upper():
                hedef_sayfa = sayfa
                break

        if not hedef_sayfa:
            print("HATA: İçinde 'ROTOR SACI' geçen bir sayfa bulunamadı!")
            return

        # Bulunan hedef_sayfa değişkenini kullanarak okuma yapıyoruz
        df = pd.read_excel(excel_dosyasi, sheet_name=hedef_sayfa, header=7, dtype=str).fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]

        # AÇIKLAMA 2 SÜTUNUNU EĞİTİME DAHİL ETTİK
        # DİKKAT: Excel'deki başlık tam olarak "AÇIKLAMA2" mi yoksa "AÇIKLAMA 2" mi kontrol et!
        girdi_sutunlari = ["GRUP", "AÇIKLAMA", "AÇIKLAMA2"]
        hedef_sutun_sira = "AÇIKLAMA S.NO"

        for sutun in girdi_sutunlari + [hedef_sutun_sira]:
            if sutun not in df.columns:
                print(f"KRİTİK HATA: '{sutun}' sütunu Excel'de yok!")
                print("Lütfen başlığı Excel ile aynı yapın. Mevcut başlıklar:", df.columns.tolist())
                return

        print("Veriler etiketleniyor...")

        encoders = {}
        X = pd.DataFrame()

        for sutun in girdi_sutunlari:
            le = LabelEncoder()
            veriler = df[sutun].astype(str).str.strip().str.upper()
            le.fit(veriler.tolist() + ["<UNK>"])
            X[sutun] = le.transform(veriler)
            encoders[sutun] = le

        y_sira = df[hedef_sutun_sira].astype(str).str.strip()
        le_sira = LabelEncoder()
        y_sira_encoded = le_sira.fit_transform(y_sira)

        print("Modeller eğitiliyor...")
        rf_sira = RandomForestClassifier(n_estimators=100, random_state=42)
        rf_sira.fit(X, y_sira_encoded)

        # ---------------- BAŞARI ORANI HESAPLAMA ----------------
        sira_tahmin = rf_sira.predict(X)
        sira_basarisi = accuracy_score(y_sira_encoded, sira_tahmin) * 100

        print("\n" + "=" * 30)
        print("🎯 YAPAY ZEKA BAŞARI ORANLARI (ROTOR SACI)")
        print(f"Sıra No Öğrenme Oranı  : %{sira_basarisi:.2f}")
        print("=" * 30 + "\n")
        # --------------------------------------------------------

        kayit_klasoru = os.path.dirname(__file__)
        joblib.dump(rf_sira, os.path.join(kayit_klasoru, "rs_sira_model.pkl"))
        joblib.dump(le_sira, os.path.join(kayit_klasoru, "rs_sira_encoder.pkl"))
        joblib.dump(encoders, os.path.join(kayit_klasoru, "rs_girdi_encoders.pkl"))

        print(f"EĞİTİM BAŞARILI! Modeller '{kayit_klasoru}' içine kaydedildi.")

    except Exception as e:
        print("Eğitim sırasında hata:")
        traceback.print_exc()

if __name__ == "__main__":
    yapay_zeka_egit_rotor_saci()