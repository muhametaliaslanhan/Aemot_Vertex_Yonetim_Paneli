from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QLabel, QLineEdit, QMessageBox, QHeaderView,QApplication)
from PyQt6.QtCore import Qt
import pandas as pd
import joblib
import os
from arayuz.mantik.json_islemleri import urun_agaci_excel_yolunu_getir

# Modelin kaydedildiği klasör (rf_model_egitimi.py ile aynı)
MODEL_KLASORU = r"C:\Users\aemot.stajyer\Desktop\AEMOT_VERTEX_YonetimPaneli\arayuz\urun_agaci_icerik_egitim"

# Dağıtım işlemleri için sözlük eşleştirmesi
PARCA_ESLESTIRME_SOZLUGU = {
    "sayfa_rotorlu_mil": ["ROTORLU", "MİL"],
    "sayfa_mil": ["MİL", "MIL"],
    "sayfa_rotor_saci": ["ROTOR SACI"],
    "sayfa_rotor_paket": ["ROTOR", "PAKET"],
    "sayfa_stator_saci": ["STATOR SACI"],
    "sayfa_stator_paket": ["STATOR", "PAKET"],
    "sayfa_bobinli_stator": ["BOBİNLİ", "STATOR"],
    "sayfa_sarilmis_bobin": ["SARILMIŞ", "BOBİN", "SARILMIS"],
    "sayfa_statorlu_govde": ["STATORLU", "GÖVDE"],
    "sayfa_govde": ["GÖVDE", "GOVDE"],
    "sayfa_govde_ayak": ["AYAK"],
    "sayfa_govde_ayak_takim": ["TAKIM", "AYAK TAKIM"],
    "sayfa_stator_rotor_pul": ["PUL"],
    "sayfa_klemens_kutu_kapagi": ["KLEMENS", "KUTU", "KAPAK"],
    "sayfa_flans": ["FLANŞ", "FLANS"],
    "sayfa_kapak": ["KAPAK"],
    "sayfa_rulman_yag_kapagi": ["RULMAN", "YAĞ", "KAPAK"],
    "sayfa_fan_muhafaza": ["FAN", "MUHAFAZA"],
    "sayfa_kanopi": ["KANOPI", "KANOPİ"],
    "sayfa_leroy_somer": ["LEROY", "SOMER"],
    "sayfa_agraf_saci": ["AGRAF", "SACI"]
}


class UrunAgaciSekmesi(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_pencere = parent  # Ana uygulamaya (main.py) dönmek için parent'ı saklıyoruz
        self.arayuzu_kur()

    def arayuzu_kur(self):
        ana_layout = QVBoxLayout()

        # ---------------- 1. ÜST BUTONLAR ----------------
        ust_buton_duzeni = QHBoxLayout()

        self.btn_ana_menu = QPushButton("Ana Menü")
        self.btn_ana_menu.setStyleSheet("background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_ana_menu.clicked.connect(self.ana_menuye_don)

        self.btn_temizle = QPushButton("Formu Sıfırla")
        self.btn_temizle.setStyleSheet("background-color: #95A5A6; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_temizle.clicked.connect(self.formu_temizle)

        # --- YAPAY ZEKA UYARI BUTONU ---
        self.btn_yz_uyari = QPushButton("⚠️ Model Bilgilendirmesi")
        self.btn_yz_uyari.setStyleSheet("""
                            background-color: #F39C12; 
                            color: white; 
                            font-weight: bold; 
                            padding: 8px 15px; 
                            border-radius: 4px;
                        """)
        self.btn_yz_uyari.setToolTip("Bu sistemin nasıl çalıştığını ve kısıtlamalarını okumak için tıklayın.")
        self.btn_yz_uyari.clicked.connect(self.uyari_bilgisi_goster)

        ust_buton_duzeni.addWidget(self.btn_yz_uyari)

        ust_buton_duzeni.addWidget(self.btn_ana_menu)
        ust_buton_duzeni.addWidget(self.btn_temizle)
        ust_buton_duzeni.addStretch()

        ana_layout.addLayout(ust_buton_duzeni)



        # NOT: Bu butonu üst layout'a (örneğin self.ust_duzen.addWidget(self.btn_yz_uyari)) eklemeyi unutma!


        # ---------------- 2. KULLANICI GİRİŞ ALANI ----------------
        form_layout = QHBoxLayout()

        self.lbl_motor_kodu = QLabel("Motor Kodu (Tepekodu):")
        self.txt_motor_kodu = QLineEdit()
        self.txt_motor_kodu.setPlaceholderText("Örn: MN31-071M020A0000000-0000")

        # YENİ: AÇIKLAMA KUTUSU
        self.lbl_motor_aciklama = QLabel("Motor Açıklaması:")
        self.txt_motor_aciklama = QLineEdit()
        self.txt_motor_aciklama.setPlaceholderText("Örn: 3 FAZLI ASENKRON MOTOR")

        self.btn_tahmin_et = QPushButton("Ürün Ağacını Oluştur")
        self.btn_tahmin_et.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_tahmin_et.clicked.connect(self.urun_agaci_olustur)

        form_layout.addWidget(self.lbl_motor_kodu)
        form_layout.addWidget(self.txt_motor_kodu)
        form_layout.addWidget(self.lbl_motor_aciklama)
        form_layout.addWidget(self.txt_motor_aciklama)
        form_layout.addWidget(self.btn_tahmin_et)

        ana_layout.addLayout(form_layout)

        # ---------------- 3. YENİ AKTARMA BUTONLARI ----------------
        aktarma_layout = QHBoxLayout()

        self.btn_tablolara_aktar = QPushButton("Tablolara Aktar (Dağıt)")
        self.btn_tablolara_aktar.setStyleSheet("background-color: #2980B9; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_tablolara_aktar.clicked.connect(self.tablolara_dagit)

        self.btn_excele_aktar = QPushButton("Toplu Excel'e Aktar")
        self.btn_excele_aktar.setStyleSheet("background-color: #8E44AD; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_excele_aktar.clicked.connect(self.excele_aktar)

        aktarma_layout.addWidget(self.btn_tablolara_aktar)
        aktarma_layout.addWidget(self.btn_excele_aktar)
        aktarma_layout.addStretch()

        ana_layout.addLayout(aktarma_layout)

        # ---------------- 4. ÜRÜN AĞACI TABLOSU ----------------
        self.tablo_urun_agaci = QTableWidget()
        self.tablo_urun_agaci.setColumnCount(9)
        self.tablo_urun_agaci.setHorizontalHeaderLabels([
            "Level",
            "Tepekodu1",
            "Tepekodu",
            "Tepekodu Açıklama",
            "Çalışma Stn",
            "Bileşen Malzeme Kodu",
            "Bileşen Malzeme Açıklama",
            "Montaj Başına (Miktar)",
            "Sıra No"
        ])
        self.tablo_urun_agaci.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.tablo_urun_agaci.horizontalHeader().setStretchLastSection(True)

        ana_layout.addWidget(self.tablo_urun_agaci)

        self.setLayout(ana_layout)

    # ---------------- FONKSİYONLAR ----------------
    def ana_menuye_don(self):
        if self.parent_pencere and hasattr(self.parent_pencere, 'sayfa_degistir'):
            self.parent_pencere.sayfa_degistir(0)

    def formu_temizle(self):
        self.txt_motor_kodu.clear()
        self.tablo_urun_agaci.setRowCount(0)

    def uyari_bilgisi_goster(self):
        mesaj = QMessageBox(self)
        mesaj.setIcon(QMessageBox.Icon.Information)
        mesaj.setWindowTitle("Yapay Zeka Modeli Hakkında Bilgilendirme")
        mesaj.setText(
            "<h3 style='color: #D35400;'>Bu Sistem Neden %100 Kesin Sonuç Vermeyebilir?</h3>"
            "<p>Bu modül, katı kurallı bir ERP varyant sistemi yerine <b>Makine Öğrenimi (Random Forest)</b> altyapısı kullanmaktadır.</p>"
            "<ul>"
            "<li>Sistem, daha önce AEMOT veritabanında olan motorları tanıyıp şablonlarını kusursuzca getirecek şekilde eğitilmiştir.</li>"
            "<li>Ancak veritabanında <b>daha önce hiç olmayan</b> yeni bir motor kodu girildiğinde, model elindeki özelliklere bakarak en benzer motoru <i>tahmin eder</i> ve onun ürün ağacını kopyalar.</li>"
            "<li>Yapay zeka, o motor için özel olarak istenmiş spesifik bir flanşı, manuel mühendislik değişikliklerini veya standart dışı özel rulman/boya taleplerini bilemez.</li>"
            "</ul>"
            "<p><b>Sonuç:</b> Yeni oluşturulan ürün ağaçlarını üretime vermeden veya onaylamadan önce mutlaka <b>gözden geçiriniz</b>.</p>"
        )
        mesaj.setStyleSheet("""
            QMessageBox { background-color: #F8F9F9; }
            QLabel { font-size: 13px; color: #2C3E50; }
            QPushButton { background-color: #34495E; color: white; padding: 6px 15px; font-weight: bold; border-radius: 3px; }
            QPushButton:hover { background-color: #2C3E50; }
        """)
        mesaj.exec()

    def urun_agaci_olustur(self):
        tepekodu = self.txt_motor_kodu.text().strip()
        aciklama_girisi = self.txt_motor_aciklama.text().strip()

        if not tepekodu or not aciklama_girisi:
            QMessageBox.warning(self, "HATA", "Lütfen Motor Kodu ve Açıklamasını eksiksiz girin!")
            return

        model_yolu = os.path.join(MODEL_KLASORU, 'rf_model.pkl')
        mapping_yolu = os.path.join(MODEL_KLASORU, 'rf_profil_mapping.pkl')

        if not os.path.exists(model_yolu) or not os.path.exists(mapping_yolu):
            QMessageBox.critical(self, "HATA", "Model veya Mapping dosyası bulunamadı! Lütfen önce eğitimi tamamlayın.")
            return

        try:
            # Modeli ve Haritayı Yükle
            rf_model = joblib.load(model_yolu)
            profil_mapping = joblib.load(mapping_yolu)

            # Encoder'ları yükle
            encoderlar = {}
            kategorik_kolonlar = ['Seri', 'Kasa_Tipi', 'Govde_Boyu', 'Kutup', 'Paket', 'Yapi_Sekli', 'Flans',
                                  'Voltaj_Frekans', 'Opsiyon_1', 'Opsiyon_2', 'Aciklama']
            for kolon in kategorik_kolonlar:
                le_path = os.path.join(MODEL_KLASORU, f"le_{kolon.lower()}.pkl")
                encoderlar[kolon] = joblib.load(le_path)

            # Tepekodu'ndan özellikleri ayrıştır
            tepekodu_str = str(tepekodu)
            seri = tepekodu_str[0:4] if len(tepekodu_str) > 4 else "UNK"
            kasa = tepekodu_str[5:8] if len(tepekodu_str) > 16 else "UNK"
            govde = tepekodu_str[8] if len(tepekodu_str) > 16 else "U"
            kutup = tepekodu_str[9:12] if len(tepekodu_str) > 16 else "UNK"
            paket = tepekodu_str[12] if len(tepekodu_str) > 16 else "U"
            yapi = tepekodu_str[13] if len(tepekodu_str) > 16 else "U"
            flans = tepekodu_str[14:16] if len(tepekodu_str) > 16 else "UN"
            voltaj = tepekodu_str[16:18] if len(tepekodu_str) > 18 else "UN"
            ops_1 = tepekodu_str[18:22] if len(tepekodu_str) > 21 else "UNK"
            ops_2 = tepekodu_str[22:] if len(tepekodu_str) > 22 else "UNK"

            # Özellikleri DataFrame'e dönüştür ve Encoder ile encode et
            ornek_df = pd.DataFrame(
                [[seri, kasa, govde, kutup, paket, yapi, flans, voltaj, ops_1, ops_2, aciklama_girisi]],
                columns=kategorik_kolonlar)
            for kolon in kategorik_kolonlar:
                le = encoderlar[kolon]
                val = str(ornek_df[kolon].iloc[0])
                if val not in le.classes_:
                    val = "UNK" if val != "U" else "U"
                ornek_df[f'{kolon}_Kod'] = le.transform([val])

            X_yeni = ornek_df[[f"{k}_Kod" for k in kategorik_kolonlar]].astype('int16')

            # Profili Tahmin Et
            tahmin_edilen_profil_id = rf_model.predict(X_yeni)[0]
            if tahmin_edilen_profil_id not in profil_mapping:
                QMessageBox.warning(self, "HATA", "Tahmin edilen profil ID haritada bulunamadı.")
                return

            bom_df = profil_mapping[tahmin_edilen_profil_id]

            # Tabloya Aktar
            self.tablo_urun_agaci.setRowCount(0)
            self.tablo_urun_agaci.setRowCount(len(bom_df))

            for row_idx, (_, row) in enumerate(bom_df.iterrows()):
                # --- NAN (BOŞ HÜCRE) VE EXCEL TARİH BUG'I TEMİZLİĞİ ---
                calisma_stn = str(row.get('Çalışma Stn', ''))
                if calisma_stn.lower() == 'nan':
                    calisma_stn = ""

                ham_miktar = row.get('Montaj Basına', 0)
                try:
                    m_float = float(ham_miktar)
                    # Tam sayıysa (1.0 -> 1)
                    if m_float.is_integer():
                        miktar_metni = str(int(m_float))
                    # Ondalıklıysa virgüllü yap (0.003 -> 0,003)
                    else:
                        miktar_metni = str(round(m_float, 4)).replace('.', ',')
                except:
                    miktar_metni = str(ham_miktar)

                self.tablo_urun_agaci.setItem(row_idx, 0, QTableWidgetItem(str(row.get('Level', ''))))
                self.tablo_urun_agaci.setItem(row_idx, 1, QTableWidgetItem(tepekodu))
                self.tablo_urun_agaci.setItem(row_idx, 2, QTableWidgetItem(str(row.get('Tepekodu', ''))))
                self.tablo_urun_agaci.setItem(row_idx, 3, QTableWidgetItem(aciklama_girisi))
                self.tablo_urun_agaci.setItem(row_idx, 4, QTableWidgetItem(calisma_stn))
                self.tablo_urun_agaci.setItem(row_idx, 5, QTableWidgetItem(str(row.get('Bılesen Malzeme Kodu', ''))))
                self.tablo_urun_agaci.setItem(row_idx, 6,
                                              QTableWidgetItem(str(row.get('Bılesen Malzeme Acıklama', ''))))
                self.tablo_urun_agaci.setItem(row_idx, 7, QTableWidgetItem(miktar_metni))
                self.tablo_urun_agaci.setItem(row_idx, 8, QTableWidgetItem(str(row.get('Sıra No', ''))))

            QMessageBox.information(self, "BAŞARILI",
                                    f"Ürün ağacı '{tahmin_edilen_profil_id}' ID'li profile göre başarıyla oluşturuldu!")

        except Exception as e:
            QMessageBox.critical(self, "HATA", f"Tahmin sırasında bir hata oluştu:\n{str(e)}")

    def tablolara_dagit(self):
        if self.tablo_urun_agaci.rowCount() == 0:
            QMessageBox.warning(self, "UYARI", "Tablo boş! Önce bir ürün ağacı oluşturun.")
            return

        if not self.parent_pencere or not hasattr(self.parent_pencere, "sayfa_motor"):
            QMessageBox.warning(self, "HATA", "Ana üretim paneline ulaşılamıyor.")
            return

        uretim_paneli = self.parent_pencere.sayfa_motor
        basarili_sayisi = 0

        for r in range(self.tablo_urun_agaci.rowCount()):
            aciklama = self.tablo_urun_agaci.item(r, 6).text().upper()
            kod = self.tablo_urun_agaci.item(r, 5).text().upper()
            
            # Dictionary matching rule
            hedef_sayfa_attr = None
            for attr_name, anahtar_kelimeler in PARCA_ESLESTIRME_SOZLUGU.items():
                if any(kelime in aciklama for kelime in anahtar_kelimeler) or any(kelime in kod for kelime in anahtar_kelimeler):
                    hedef_sayfa_attr = attr_name
                    break
            
            # If found, add row to the target table
            if hedef_sayfa_attr and hasattr(uretim_paneli, hedef_sayfa_attr):
                hedef_sekme = getattr(uretim_paneli, hedef_sayfa_attr)
                if hasattr(hedef_sekme, "tablo"):
                    hedef_tablo = hedef_sekme.tablo
                    mevcut_satir = hedef_tablo.rowCount()
                    
                    # Temporarily disable sorting before inserting
                    sorting_enabled = hedef_tablo.isSortingEnabled()
                    hedef_tablo.setSortingEnabled(False)
                    
                    hedef_tablo.insertRow(mevcut_satir)
                    # Tablolara uygun formatta (genellikle 0=KOD, 1=AÇIKLAMA vb.) veri gönderiyoruz
                    # Parça tablolarının genel yapısı: #KOD, AÇIKLAMA ... 
                    hedef_tablo.setItem(mevcut_satir, 0, QTableWidgetItem(kod))
                    hedef_tablo.setItem(mevcut_satir, 1, QTableWidgetItem(aciklama))
                    
                    # Eğer Miktar veya başka kolon eklenecekse burası genişletilebilir
                    if hedef_tablo.columnCount() > 2:
                        miktar = self.tablo_urun_agaci.item(r, 7).text()
                        hedef_tablo.setItem(mevcut_satir, 2, QTableWidgetItem(miktar))

                    hedef_tablo.setSortingEnabled(sorting_enabled)
                    basarili_sayisi += 1

        QMessageBox.information(self, "DAĞITIM TAMAMLANDI", f"Toplam {basarili_sayisi} adet bileşen ilgili tablolara aktarıldı.")

    def excele_aktar(self):
        if self.tablo_urun_agaci.rowCount() == 0:
            QMessageBox.warning(self, "UYARI", "Tablo boş! Aktarılacak veri yok.")
            return

        excel_yolu = urun_agaci_excel_yolunu_getir()
        if not excel_yolu or not (excel_yolu.endswith(".xlsx") or excel_yolu.endswith(".xls")):
            QMessageBox.warning(self, "HATA", "Geçerli bir Excel kayıt yolu bulunamadı!\nAyarlar menüsünden Ürün Ağacı Kayıt Exceli yolunu belirleyin.")
            return

        veriler = []
        basliklar = [self.tablo_urun_agaci.horizontalHeaderItem(i).text() for i in range(self.tablo_urun_agaci.columnCount())]

        for r in range(self.tablo_urun_agaci.rowCount()):
            satir = []
            for c in range(self.tablo_urun_agaci.columnCount()):
                item = self.tablo_urun_agaci.item(r, c)
                satir.append(item.text() if item else "")
            veriler.append(satir)

        df_yeni = pd.DataFrame(veriler, columns=basliklar)

        try:
            if os.path.exists(excel_yolu):
                df_eski = pd.read_excel(excel_yolu)
                df_son = pd.concat([df_eski, df_yeni], ignore_index=True)
            else:
                df_son = df_yeni
            
            df_son.to_excel(excel_yolu, index=False)
            QMessageBox.information(self, "BAŞARILI", f"Veriler başarıyla excel dosyasına kaydedildi!\n\nYol: {excel_yolu}")
        except Exception as e:
            QMessageBox.critical(self, "HATA", f"Excel'e kaydedilirken bir hata oluştu (Dosya açık olabilir):\n{str(e)}")

    def keyPressEvent(self, event):
        # Ctrl + C basıldığında çalışacak
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_C:
            secili_alan = self.tablo_urun_agaci.selectedRanges()
            if not secili_alan:
                return

            kopya_metin = ""
            for r in range(secili_alan[0].topRow(), secili_alan[0].bottomRow() + 1):
                satir = []
                for c in range(secili_alan[0].leftColumn(), self.tablo_urun_agaci.columnCount()):
                    item = self.tablo_urun_agaci.item(r, c)
                    # Sadece seçili hücreleri almak için
                    if item and item.isSelected():
                        satir.append(item.text())
                if satir:
                    kopya_metin += "\t".join(satir) + "\n"

            # Kopyalanan metni panoya (clipboard) gönder
            QApplication.clipboard().setText(kopya_metin)
        else:
            # Diğer tuş kombinasyonlarının normal çalışması için
            super().keyPressEvent(event)

if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    pencere = UrunAgaciSekmesi()
    pencere.show()
    sys.exit(app.exec())