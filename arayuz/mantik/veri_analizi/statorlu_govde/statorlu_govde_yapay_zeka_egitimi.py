import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import accuracy_score  # BAŞARI ORANI İÇİN EKLENDİ
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


def yapay_zeka_egit_statorlu_govde():
    print("STATORLU GOVDE eğitimi başlatılıyor...")

    excel_yolu = excel_yolunu_bul()

    if not os.path.exists(excel_yolu):
        print("HATA: Excel dosyası bulunamadı!")
        return

    try:
        excel_dosyasi = pd.ExcelFile(excel_yolu)
        hedef_sayfa = None

        for sayfa in excel_dosyasi.sheet_names:
            if "STATORLU" in sayfa.upper() and "GOVDE" in sayfa.upper():
                hedef_sayfa = sayfa
                break

        if not hedef_sayfa:
            print("HATA: İçinde 'STATORLU' veya 'GOVDE' geçen bir sayfa bulunamadı!")
            return

        df = pd.read_excel(excel_dosyasi, sheet_name=hedef_sayfa, header=8, dtype=str).fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]

        girdi_sutunlari = ["GRUP", "AÇIKLAMA"]
        hedef_sutun_resim = "RESİM NO"
        hedef_sutun_sira = "AÇIKLAMA S.NO"

        for sutun in girdi_sutunlari + [hedef_sutun_resim, hedef_sutun_sira]:
            if sutun not in df.columns:
                print(f"KRİTİK HATA: '{sutun}' sütunu Excel'de yok!")
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

        y_resim = df[hedef_sutun_resim].astype(str).str.strip()
        y_sira = df[hedef_sutun_sira].astype(str).str.strip()

        le_resim = LabelEncoder()
        le_sira = LabelEncoder()

        y_resim_encoded = le_resim.fit_transform(y_resim)
        y_sira_encoded = le_sira.fit_transform(y_sira)

        print("Modeller eğitiliyor...")
        rf_resim = RandomForestClassifier(n_estimators=100, random_state=42)
        rf_resim.fit(X, y_resim_encoded)

        rf_sira = RandomForestClassifier(n_estimators=100, random_state=42)
        rf_sira.fit(X, y_sira_encoded)

        # ---------------- BAŞARI ORANI HESAPLAMA (BELLEK DOSTU YÖNTEM) ----------------
        print("Başarı oranı hesaplanıyor (RAM şişmesin diye parça parça tahmin ediliyor)...")

        def parcali_tahmin(model, veri, parca_boyutu=1000):
            tahminler = []
            for i in range(0, len(veri), parca_boyutu):
                parca = veri.iloc[i:i + parca_boyutu]
                tahminler.extend(model.predict(parca))
            return tahminler

        resim_tahmin = parcali_tahmin(rf_resim, X)
        sira_tahmin = parcali_tahmin(rf_sira, X)

        resim_basarisi = accuracy_score(y_resim_encoded, resim_tahmin) * 100
        sira_basarisi = accuracy_score(y_sira_encoded, sira_tahmin) * 100

        print("\n" + "=" * 30)
        print("🎯 YAPAY ZEKA BAŞARI ORANLARI (STATORLU GOVDE)")
        print(f"Resim No Öğrenme Oranı : %{resim_basarisi:.2f}")
        print(f"Sıra No Öğrenme Oranı  : %{sira_basarisi:.2f}")
        print("=" * 30 + "\n")
        # --------------------------------------------------------

        kayit_klasoru = os.path.dirname(__file__)
        joblib.dump(rf_resim, os.path.join(kayit_klasoru, "rs_resim_model.pkl"))
        joblib.dump(le_resim, os.path.join(kayit_klasoru, "rs_resim_encoder.pkl"))
        joblib.dump(rf_sira, os.path.join(kayit_klasoru, "rs_sira_model.pkl"))
        joblib.dump(le_sira, os.path.join(kayit_klasoru, "rs_sira_encoder.pkl"))
        joblib.dump(encoders, os.path.join(kayit_klasoru, "rs_girdi_encoders.pkl"))

        print(f"EĞİTİM BAŞARILI! Modeller kaydedildi.")

    except Exception as e:
        print("Eğitim sırasında hata:")
        traceback.print_exc()


if __name__ == "__main__":
    yapay_zeka_egit_statorlu_govde()