import joblib
import pandas as pd
import os

# Modellerin bulunduğu klasör (pkl dosyaları buradaysa sorun yok, farklıysa yolu düzenle)
MANTIK_KLASORU = os.path.dirname(__file__)
MODEL_YOLU = os.path.join(MANTIK_KLASORU, "veri_analizi")  # pkl dosyaların veri_analizi içindeyse

try:
    # Modelleri RAM'e yükle (Sistem açıldığında 1 kere yüklenir, hızlı çalışır)
    model = joblib.load(os.path.join(MODEL_YOLU, 'sira_no_modeli.pkl'))
    vektorizer = joblib.load(os.path.join(MODEL_YOLU, 'tfidf_vektorizer.pkl'))
    egitim_sutunlari = joblib.load(os.path.join(MODEL_YOLU, 'egitim_sutunlari.pkl'))
except Exception as e:
    print(f"Yapay zeka modelleri yüklenemedi! Hata: {e}")


def yz_sira_no_tahmin_et(verimlilik, grup, kutup, yapi, flans, malzeme, ek_aciklama):
    try:
        # 1. İçi tamamen 0.0 dolu (Ondalıklı matrisleri kabul etmesi için)
        tek_satir = pd.DataFrame(0.0, index=[0], columns=egitim_sutunlari)

        # 2. Ek açıklamayı sayısallaştır ve ilgili sütunlara yerleştir
        if ek_aciklama:
            metin_matrisi = vektorizer.transform([ek_aciklama]).toarray()[0]
            metin_sutun_isimleri = vektorizer.get_feature_names_out()

            for i, sutun_adi in enumerate(metin_sutun_isimleri):
                if sutun_adi in egitim_sutunlari:
                    tek_satir.at[0, sutun_adi] = metin_matrisi[i]

        # 3. Kategorik seçimleri modele uygun şekilde "1" yap
        kategorik_ozellikler = {
            "VERİMLİLİK": verimlilik,
            "GRUP": grup,
            "KUTUP": kutup,
            "YAPI ŞEKLİ": yapi,
            "FLANŞ ÖLÇÜSÜ": flans,
            "GÖVDE MALZEMESİ": malzeme
        }

        for baslik, deger in kategorik_ozellikler.items():
            hedef_sutun = f"{baslik}_{deger}"
            # Eğer adamın arayüzde seçtiği şey modelin eğitiminde varsa onu 1 yap
            if hedef_sutun in egitim_sutunlari:
                tek_satir.at[0, hedef_sutun] = 1

        # 4. Yapay zekaya tahmini yaptır
        tahmin = model.predict(tek_satir)[0]
        return str(tahmin)

    except Exception as e:
        print(f"YZ Tahmin Hatası: {e}")
        return "0000"  # Model patlarsa varsayılan olarak 0000 dönsün sistemi kitlemesin