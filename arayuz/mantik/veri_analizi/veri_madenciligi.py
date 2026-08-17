import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score
import joblib

print("Veri yükleniyor ve temizleniyor, lütfen bekleyin...")

# 1. Veri Okuma (Başlıklar 12. satırda olduğu için header=11)
dosya_yolu = r'C:\Users\aemot.stajyer\Desktop\Muhammetali\KODOLUŞTURUCU2018GENELSOHEIL.xlsx'
df = pd.read_excel(dosya_yolu, header=11, dtype=str)

orijinal_satir_sayisi = len(df)

# 2. Sütun Kırpma (Sadece analize girecek cerrahi sütunlar)
hedef_sutunlar = ['VERİMLİLİK','AÇIKLAMA', 'GRUP', 'KUTUP','YAPI ŞEKLİ', 'FLANŞ ÖLÇÜSÜ', 'GÖVDE MALZEMESİ','EK AÇIKLAMA', 'AÇIKLAMA NO', 'RESİM NO']

# Bazen Excel'de görünmez boşluklar olur başlıkta, patlamamak için başlıkları temizleyelim
df.columns = df.columns.str.strip().str.upper()

# Sadece istediğimiz sütunları alıyoruz
df = df[hedef_sutunlar].copy()

# 3. Veri Standardizasyonu
df = df.fillna("") # NaN olan her yeri temiz "" (boş string) yap

for sutun in df.columns:
    # Metne çevir, büyük harf yap, sağdaki soldaki sinsi boşlukları tıraşla
    df[sutun] = df[sutun].astype(str).str.upper().str.strip()

# Fazladan iç boşlukları (örn: "400/690V    ISITICI") tek boşluğa indirgeyelim
df.replace(r'\s+', ' ', regex=True, inplace=True)

df['EK AÇIKLAMA'] = df['EK AÇIKLAMA'].str.replace("PAYLAŞIM İÇİN", "", regex=False)
df['EK AÇIKLAMA'] = df['EK AÇIKLAMA'].str.replace("PAYLASIM ICIN", "", regex=False)

# 4. Kara Liste Filtrelemesi
kara_liste = ["İPTAL", "IPTAL", "YANLIŞ", "YANLIS", "DENEME", "ESKİ"]
kara_liste_regex = '|'.join(kara_liste)

# Ne 'AÇIKLAMA' içinde ne de 'EK AÇIKLAMA' içinde bu kelimeler GEÇMESİN (~ işareti tersine çevirir)
maske = ~df['AÇIKLAMA'].str.contains(kara_liste_regex, na=False) & \
        ~df['EK AÇIKLAMA'].str.contains(kara_liste_regex, na=False)

df_temiz = df[maske].copy()

# 5. Boş Hedeflerin Silinmesi
# Resim No veya Açıklama No'su tamamen boş olanları at
df_temiz = df_temiz[(df_temiz['RESİM NO'] != "") & (df_temiz['AÇIKLAMA NO'] != "")]

temiz_satir_sayisi = len(df_temiz)

print("-" * 40)
print(f"Orijinal Veri : {orijinal_satir_sayisi} satır.")
print(f"Eğitime Girecek Veri: {temiz_satir_sayisi} satır.")
print("-" * 40)

# Hedef Değişkeni Ayır
y = df_temiz['AÇIKLAMA NO'].reset_index(drop=True)

# Kategorik Özellikleri 0-1 Matrisine Çevir (One-Hot Encoding)
kategorik_X = pd.get_dummies(df_temiz[['VERİMLİLİK', 'GRUP', 'KUTUP', 'YAPI ŞEKLİ', 'FLANŞ ÖLÇÜSÜ', 'GÖVDE MALZEMESİ']], dtype=int)
kategorik_X = kategorik_X.reset_index(drop=True)

# Ek Açıklamayı Vektörlere Çevir (TF-IDF NLP İşlemi)
vektorizer = TfidfVectorizer(max_features=500, ngram_range=(1, 2))
metin_matrisi = vektorizer.fit_transform(df_temiz['EK AÇIKLAMA']).toarray()
metin_sutun_isimleri = vektorizer.get_feature_names_out()

# Sayısal metin verisini DataFrame'e dönüştür
metin_X = pd.DataFrame(metin_matrisi, columns=metin_sutun_isimleri)

# Matrisleri Birleştir (Ana Eğitim Verisi X)
X = pd.concat([kategorik_X, metin_X], axis=1)

print("Veri eğitim ve test olarak ikiye ayrılıyor...")
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

print("Yapay Zeka (Random Forest) dengeli sınıf ağırlığı ile eğitiliyor...")

model = RandomForestClassifier(n_estimators=100, random_state=42, n_jobs=-1, class_weight='balanced')
model.fit(X_train, y_train)

print("Eğitim bitti! Şimdi model test ediliyor...")
y_pred = model.predict(X_test)
basari_orani = accuracy_score(y_test, y_pred)

print("-" * 40)
print(f"MODELİN DOĞRULUK (BAŞARI) ORANI: %{basari_orani * 100:.2f}")
print("-" * 40)

print("Modelin beyni dosyaya yazdırılıyor...")
joblib.dump(model, 'sira_no_modeli.pkl')
joblib.dump(vektorizer, 'tfidf_vektorizer.pkl')
joblib.dump(X.columns.tolist(), 'egitim_sutunlari.pkl')

print("HARİKA! 3 adet '.pkl' dosyası başarıyla kaydedildi.")

print("\nModelin yanlış tahmin ettiği satırlar tespit ediliyor...")

# Test verisindeki gerçek değerler ile modelin tahminlerini yan yana getiren bir tablo oluştur
test_sonuclari = pd.DataFrame({
    'GERÇEK_SIRA_NO': y_test,
    'MODEL_TAHMİNİ': y_pred
}, index=y_test.index)

# Sadece modelin çuvalladığı (gerçek değerin tahmine eşit olmadığı) satırları filtrele
yanlis_tahminler = test_sonuclari[test_sonuclari['GERÇEK_SIRA_NO'] != test_sonuclari['MODEL_TAHMİNİ']]

# Bu yanlış tahminleri, orijinal verideki özellikler ve açıklamalarla birleştir
analiz_df = yanlis_tahminler.join(df_temiz[[
    'AÇIKLAMA', 'EK AÇIKLAMA', 'GRUP', 'KUTUP', 'YAPI ŞEKLİ',
    'FLANŞ ÖLÇÜSÜ', 'VERİMLİLİK', 'GÖVDE MALZEMESİ'
]])

# Sütun sırasını incelemesi kolay olacak şekilde düzenle
analiz_df = analiz_df[[
    'GERÇEK_SIRA_NO', 'MODEL_TAHMİNİ', 'AÇIKLAMA', 'EK AÇIKLAMA',
    'GRUP', 'KUTUP', 'YAPI ŞEKLİ', 'FLANŞ ÖLÇÜSÜ', 'VERİMLİLİK', 'GÖVDE MALZEMESİ'
]]

# Çıkan sonucu Excel'e kaydet
analiz_dosyasi = 'hatali_tahminler_analizi.xlsx'
analiz_df.to_excel(analiz_dosyasi, index=False)

print("-" * 40)
print(f"TEŞHİS TAMAMLANDI! Toplam {len(analiz_df)} adet hatalı tahmin bulundu.")
print(f"Lütfen klasöründeki '{analiz_dosyasi}' dosyasını açıp incele.")
print("-" * 40)

#    def kayit_baslat(self):
#        try:
#            ham_aciklama = self.ek_aciklama_kutusu.text()
#            temiz_aciklama = turkce_karakter_temizleme(ham_aciklama).replace("  ", " ").strip()
#
#            # 1. ZORUNLU ÖZEL KW/DEVİR KONTROLÜ
#            if not self.cb_standart.isChecked():
#                has_kw = self.cb_ozel_kw.isChecked() and bool(self.le_ozel_kw.text().strip())
#                has_devir = self.cb_ozel_devir.isChecked() and bool(self.le_ozel_devir.text().strip())
#                if not (has_kw and has_devir):
#                    QMessageBox.warning(self, "EKSİK BİLGİ",
#                                        "Motor 'STANDART' seçilmediyse, 'Özel KW' ve 'Özel Devir' girmek zorundasınız!")
#                    return
#
#            # 2. YASAKLI KELİME KONTROLÜ
#            yasakli_kelimeler = ["TROPIKALIZELI", "MARIN", "MARINE", "TERMISTOR", "TERMIK", "ISITICI",
#                                 "GUDUK", "-V", "PASLANMAZ", "HZ", "BOYA", "RAL", "IP56", "IP65", "IP66", "IP67"]
#            for kelime in yasakli_kelimeler:
#                if kelime in temiz_aciklama:
#                    QMessageBox.critical(self, "KURAL İHLALİ",
#                                         f"Yasaklı Kelime: {kelime}\nBunu ek açıklama yerine onay kutularından seçiniz.")
#                    return
#
#            # 3. UYARILAR BİRLEŞTİRME (Hepsi tek seferde sorulacak)
#            uyarilar = []
#            yapi = self.yapi_sekli_kutusu.currentText()
#            flans = self.flans_olcusu_kutusu.currentText()
#
#            if yapi in ["B3", "B9"] and flans != "YOK":
#                uyarilar.append(f"{yapi} yapı şeklinde Flanş 'YOK' olmalıdır.")
#            elif yapi not in ["B3", "B9"] and flans == "YOK":
#                uyarilar.append(f"{yapi} yapı şeklinde Flanş GİRİLMELİDİR.")
#
#            try:
#                m_degeri = int(self.M_kutusu.currentText())
#            except:
#                m_degeri = 0
#            malzeme = self.govde_malzemesi_kutusu.currentText()
#
#            if 63 <= m_degeri <= 112 and malzeme != "ALUMINYUM":
#                uyarilar.append(f"{m_degeri} gövde standart ALÜMİNYUM olmalıdır.")
#            elif 225 <= m_degeri <= 400 and malzeme != "PIK":
#                uyarilar.append(f"{m_degeri} gövde standart PIK olmalıdır.")
#            elif m_degeri > 400 and malzeme != "CELIK":
#                uyarilar.append(f"{m_degeri} gövde standart ÇELİK olmalıdır.")
#            if m_degeri >= 200 and malzeme == "ALUMINYUM":
#                uyarilar.append(f"{m_degeri} gövdede ALÜMİNYUM üretilmez!")
#
#            if uyarilar:
#                uyari_metni = "\n".join([f"• {u}" for u in uyarilar])
#                cevap = QMessageBox.question(self, "STANDART DIŞI ONAYI",
#                                             f"Standart dışı seçimler tespit edildi:\n\n{uyari_metni}\n\nYine de kod üretmek istiyor musunuz?",
#                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
#                if cevap == QMessageBox.StandardButton.No: return
#
#            # 4. IP, MARINE VE TERMİSTÖR KURALLARI
#            ip_degeri = self.ip_sinifi_kutusu.currentText()
#            marine_secili = self.cb_marine.isChecked()
#            termistor_secili = self.cb_termistor.isChecked()
#
#            if marine_secili:
#                if ip_degeri == "56": ip_degeri = ""  # Marine zaten IP56'dır
#                termistor_secili = False  # Marine motor zaten termistörlüdür
#            if m_degeri >= 200:
#                termistor_secili = False  # Büyük motorlarda termistör yazılmaz
#
#            # 5. EK AÇIKLAMA BÖLÜNÜYOR (Çift Yazmayı Engelliyoruz!)
#            yeni_eklemeler = []
#            if self.cb_cd.isChecked(): yeni_eklemeler.append("CD")
#            if self.cb_onden_sabit.isChecked(): yeni_eklemeler.append("ONDEN SABIT")
#            if self.cb_enmotor.isChecked(): yeni_eklemeler.append("ENMOTOR")
#            if self.cb_imak.isChecked(): yeni_eklemeler.append("IMAK")
#            if self.cb_regl.isChecked(): yeni_eklemeler.append("REGL")
#            if self.cb_teco.isChecked(): yeni_eklemeler.append("TECO")
#
#            temiz_boya = turkce_karakter_temizleme(self.le_boya.text())
#            if self.cb_boya.isChecked() and temiz_boya:
#                yeni_eklemeler.append(f"{temiz_boya} BOYALI")
#
#            # Excel'deki EK AÇIKLAMA sütununa yazılacak saf metin
#            gercek_ek_aciklama = f"{temiz_aciklama} {' '.join(yeni_eklemeler)}".strip()
#
#            # Yapay Zeka'ya gidecek HARMANLANMIŞ ve ALFABETİK metin
#            ai_kelimeler = gercek_ek_aciklama.split()
#            if self.cb_isitici.isChecked(): ai_kelimeler.append("ISITICI")
#            if self.cb_termik.isChecked(): ai_kelimeler.append("TERMIK")
#            if marine_secili: ai_kelimeler.append("MARINE")
#            if termistor_secili: ai_kelimeler.append("TERMISTOR")
#
#            ai_kelimeler.sort()
#            ai_icin_ek_aciklama = " ".join(ai_kelimeler)
#
#            # 6. SIFIR KOPYA (AYNI MOTOR KONTROLÜ)
#            kontrol_ozellikleri = {
#                'VERİMLİLİK': self.verimlilik_kutusu.currentText(),
#                'GRUP': self.M_kutusu.currentText(),
#                'GÖVDE BOYU': self.govde_boyu_kutusu.currentText(),
#                'KUTUP': self.kutup_kutusu.currentText(),
#                'PAKET BOYU': self.paket_boyu_kutusu.currentText(),
#                'YAPI ŞEKLİ': self.yapi_sekli_kutusu.currentText(),
#                'FLANŞ ÖLÇÜSÜ': self.flans_olcusu_kutusu.currentText(),
#                'GÖVDE MALZEMESİ': malzeme,
#                'EK AÇIKLAMA': gercek_ek_aciklama  # Kopya kontrolü saf metin üzerinden yapılır
#            }
#
#            if hasattr(self, 'excel_verisi') and self.excel_verisi is not None:
#                kopya_var_mi, kopya_satir = ayni_motor_var_mi(self.excel_verisi, kontrol_ozellikleri)
#                if kopya_var_mi:
#                    cevap = QMessageBox.question(
#                        self, "BU MOTOR ZATEN VAR!",
#                        f"Girdiğiniz özelliklerin BİREBİR AYNISI Excel'de zaten kayıtlı.\n\n"
#                        f"TİP ID: {kopya_satir.get('AÇIKLAMA KOD', kopya_satir.get('KOD', 'Bulunamadı'))}\n"
#                        f"SIRA NO: {kopya_satir.get('AÇIKLAMA NO', 'Bulunamadı')}\n\n"
#                        f"Bu motoru yeni bir kod üretmeden direkt tabloya aktarmak ister misiniz?",
#                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
#                    )
#                    if cevap == QMessageBox.StandardButton.Yes:
#                        mevcut_satir = self.tablo.rowCount()
#                        self.tablo.insertRow(mevcut_satir)
#
#                        kopya_veri = [
#                            kopya_satir.get('KOD', kopya_satir.get('#KOD', '')), kopya_satir.get('AÇIKLAMA', ''),
#                            kopya_satir.get('VERİMLİLİK', ''), kopya_satir.get('GRUP', ''),
#                            kopya_satir.get('GÖVDE BOYU', ''), kopya_satir.get('KUTUP', ''),
#                            kopya_satir.get('PAKET BOYU', ''), kopya_satir.get('YAPI ŞEKLİ', ''),
#                            kopya_satir.get('FLANŞ ÖLÇÜSÜ', ''), kopya_satir.get('GÖVDE MALZEMESİ', ''),
#                            kopya_satir.get('RESİM NO', ''), kopya_satir.get('EK AÇIKLAMA', ''),
#                            kopya_satir.get('AÇIKLAMA NO', ''), kopya_satir.get('AÇIKLAMA KOD', ''),
#                            kopya_satir.get('KW/DEVİR', ''), QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")
#                        ]
#
#                        for sutun, veri in enumerate(kopya_veri):
#                            hucre = QTableWidgetItem(str(veri))
#                            self.tablo.setItem(mevcut_satir, sutun, hucre)
#                        return
#                    else:
#                        return
#
#                        # 7. YENİ MOTOR OLUŞTURMA (Kopya değilse burası çalışır)
#            olusturulan_motor = motor(
#                versiyon=self.versiyon_kutusu.currentText(),
#                verimlilik=self.verimlilik_kutusu.currentText(),
#                M=self.M_kutusu.currentText(),
#                govde_boyu=self.govde_boyu_kutusu.currentText(),
#                kutup=self.kutup_kutusu.currentText(),
#                paket_boyu=self.paket_boyu_kutusu.currentText(),
#                yapi_sekli=self.yapi_sekli_kutusu.currentText(),
#                flans_olcusu=self.flans_olcusu_kutusu.currentText(),
#                govde_malzemesi=malzeme,
#                resim_no=self.resim_no_kutusu.currentText(),
#                sira_no="",
#                ek_aciklama=gercek_ek_aciklama,  # Marine 2 kere yazmasın diye saf hali gidiyor
#                ip_sinifi=ip_degeri,
#                ozel_frekans=self.ozel_frekans_kutusu.currentText(),
#                ozel_gerilim=turkce_karakter_temizleme(self.ozel_gerilim_kutusu.text()),
#                paslanmaz_civata=turkce_karakter_temizleme(self.paslanmaz_civata_kutusu.currentText()),
#                boya_durum=self.cb_boya.isChecked(),
#                boya_deger=temiz_boya,
#                ozel_kw_durum=self.cb_ozel_kw.isChecked(),
#                ozel_kw_deger=self.le_ozel_kw.text(),
#                ozel_devir_durum=self.cb_ozel_devir.isChecked(),
#                ozel_devir_deger=self.le_ozel_devir.text(),
#                standart=self.cb_standart.isChecked(),
#                isitici=self.cb_isitici.isChecked(),
#                marine=marine_secili,
#                guduk=self.cb_guduk.isChecked(),
#                termistor=termistor_secili,
#                vantilator=self.cb_vantilator.isChecked(),
#                termik=self.cb_termik.isChecked(),
#                paslanmaz_rakor=self.cb_paslanmaz_rakor.isChecked(),
#                tropikalizeli=self.cb_tropikalizeli.isChecked()
#            )
#
#            uretilen_aciklama = olusturulan_motor.aciklama_uret()
#
#            # 8. YAPAY ZEKA SIRA NO TAHMİNİ
#            yz_tahmini = yz_sira_no_tahmin_et(
#                verimlilik=self.verimlilik_kutusu.currentText(),
#                grup=self.M_kutusu.currentText(),
#                kutup=self.kutup_kutusu.currentText(),
#                yapi=self.yapi_sekli_kutusu.currentText(),
#                flans=self.flans_olcusu_kutusu.currentText(),
#                malzeme=malzeme,
#                ek_aciklama=ai_icin_ek_aciklama  # Yapay Zeka her şeyi bilmek zorunda
#            )
#
#            # Excel Formatına Uygun Sıra No ve Kod Düzenlemesi
#            sira_no_tam = str(yz_tahmini).zfill(4)
#            olusturulan_motor.sira_no = sira_no_tam
#
#            aciklama_no = str(int(sira_no_tam))  # Örn: "25"
#            aciklama_kod = sira_no_tam  # Örn: "0025"
#
#            uretilen_id = olusturulan_motor.tip_id_uret()
#
#            kw_devir_metni = ""
#            if self.cb_ozel_kw.isChecked() and self.cb_ozel_devir.isChecked():
#                kw_devir_metni = f"{self.le_ozel_kw.text()} kW / {self.le_ozel_devir.text()} d/dk"
#
#            bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")
#
#            yeni_kayit = [
#                uretilen_id, uretilen_aciklama, self.verimlilik_kutusu.currentText(),
#                self.M_kutusu.currentText(), self.govde_boyu_kutusu.currentText(),
#                self.kutup_kutusu.currentText(), self.paket_boyu_kutusu.currentText(),
#                self.yapi_sekli_kutusu.currentText(), self.flans_olcusu_kutusu.currentText(),
#                malzeme, self.resim_no_kutusu.currentText(),
#                gercek_ek_aciklama, aciklama_no, aciklama_kod, kw_devir_metni, bugunun_tarihi
#            ]
#
#            mevcut_satir = self.tablo.rowCount()
#            self.tablo.insertRow(mevcut_satir)
#            for sutun, veri in enumerate(yeni_kayit):
#                hucre = QTableWidgetItem(str(veri))
#                self.tablo.setItem(mevcut_satir, sutun, hucre)
#
#        except Exception as e:
#            hata_mesaji = traceback.format_exc()
#            QMessageBox.critical(self, "SİSTEM ÇÖKMEKTEN KURTARILDI",
#                                 f"Kayıt sırasında bir hata oluştu:\n\n{hata_mesaji}")