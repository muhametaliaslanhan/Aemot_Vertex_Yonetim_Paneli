from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QSpacerItem, QSizePolicy,
    QScrollArea, QMessageBox, QCompleter, QDialog, QDateEdit, QFormLayout, QMenu
)
from PyQt6.QtCore import Qt, QTimer, QDateTime,QDate
from PyQt6.QtGui import QPixmap
import pandas as pd
import traceback
import os
import stat
import shutil
import tempfile
import xlwings as xw
from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir, cap_oku, kaynak_yolu
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.motor_verileri import turkce_karakter_temizleme
from arayuz.mantik import oturum


class MilArayuzu(QWidget):
    def __init__(self, ana_pencere):
        super().__init__()
        self.ana_pencere = ana_pencere

        # Genel Arayüz Stili
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', Arial; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #BDC3C7; border-radius: 4px; background: white; }
            QLineEdit:focus, QComboBox:focus { border: 2px solid #3498DB; background-color: #EBF5FB; }
            QGroupBox { font-weight: bold; border: 1px solid #BDC3C7; border-radius: 6px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; color: #2C3E50; }
            QLabel { font-weight: bold; color: #34495E; }
        """)

        ana_duzen = QVBoxLayout(self)

        # --- YENİ: ÜST BUTONLAR ---
        ust_buton_duzeni = QHBoxLayout()
        btn_geri = QPushButton("Ana Menü")
        btn_geri.setStyleSheet(
            "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        btn_geri.clicked.connect(lambda: self.ana_pencere.sayfa_degistir(0))

        self.btn_formu_temizle = QPushButton("Formu Sıfırla")
        self.btn_formu_temizle.setStyleSheet(
            "background-color: #95A5A6; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_formu_temizle.clicked.connect(self.formu_temizle)

        ust_buton_duzeni.addWidget(btn_geri)
        ust_buton_duzeni.addWidget(self.btn_formu_temizle)
        ust_buton_duzeni.addStretch()
        ana_duzen.addLayout(ust_buton_duzeni)

        # ---------------- 1. ÜST BİLGİ ALANI (HEADER) ----------------
        header_duzen = QHBoxLayout()

        self.logo_etiketi = QLabel()
        # EXE LOGO SORUNU ÇÖZÜMÜ: kaynak_yolu eklendi!
        self.logo_etiketi.setPixmap(QPixmap(kaynak_yolu("arayuz/varliklar/aemot_logo_gorseli.png")))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)

        baslik = QLabel("MİL İD VE AÇIKLAMA OLUŞTURUCU")
        baslik.setStyleSheet("font-size: 20px; font-weight: bold; color: #2980B9;")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.saat_etiketi = QLabel("SAAT YÜKLENİYOR...")
        self.saat_etiketi.setStyleSheet("font-size: 14px; font-weight: bold; color: #2980B9;")

        header_duzen.addWidget(self.logo_etiketi)
        header_duzen.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        header_duzen.addWidget(baslik)
        header_duzen.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        header_duzen.addWidget(self.saat_etiketi)

        ana_duzen.addLayout(header_duzen)

        # ---------------- 2. FORM ALANI (Scroll Edilebilir) ----------------
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)

        form_widget = QWidget()
        form_ana_duzen = QVBoxLayout(form_widget)

        # --- GRUP 1: Temel Özellikler ---
        grup_temel = QGroupBox("Temel Özellikler (Ana Motordan Otomatik Gelecek)")
        duzen_temel = QGridLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["ESKİ_GÖVDE", "PREMIUM"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630","PILOTCAR"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(["2", "4", "6", "8", "10", "12", "2/4", "4/2","4/6"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(["S", "M", "L", "S-M", "S-L", "M-L", "S-M-L","260PK","280PK","300PK","310PK","340Pk"
                                            ,"360PK","400PK","420PK","450PK","470PK","475PK","520PK","525PK","550PK","585PK"
                                            ,"650PK","660PK","685PK","750PK","920PK","M-725PK"])

        duzen_temel.addWidget(QLabel("Versiyon:"), 0, 0)
        duzen_temel.addWidget(self.versiyon_kutusu, 0, 1)
        duzen_temel.addWidget(QLabel("Grup (M):"), 0, 2)
        duzen_temel.addWidget(self.grup_kutusu, 0, 3)

        duzen_temel.addWidget(QLabel("Kutup:"), 1, 0)
        duzen_temel.addWidget(self.kutup_kutusu, 1, 1)
        duzen_temel.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        duzen_temel.addWidget(self.govde_boyu_kutusu, 1, 3)

        grup_temel.setLayout(duzen_temel)
        form_ana_duzen.addWidget(grup_temel)

        # --- GRUP 2: Mile Özel Girdiler ---
        grup_ozel = QGroupBox("Mil Detayları")
        duzen_ozel = QGridLayout()

        self.kesilmis_islenmis=QComboBox()
        self.kesilmis_islenmis.addItems(["KESILMIS","ISLENMIS"])

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        self.resim_no_input = QLineEdit()
        self.resim_no_input.setPlaceholderText("Örn: 3130")

        self.kesilmis_mil_boyu =QLineEdit()
        self.mil_capi=QLineEdit()

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Örn: BAKIR BARALI")

        duzen_ozel.addWidget(QLabel("Kesilmiş/İşlenmiş:"), 0, 0)
        duzen_ozel.addWidget(self.kesilmis_islenmis, 0, 1)
        duzen_ozel.addWidget(QLabel("Revizyon No:"), 0, 2)
        duzen_ozel.addWidget(self.revizyon_kutusu, 0, 3)

        duzen_ozel.addWidget(QLabel("Resim No:"), 1, 0)
        duzen_ozel.addWidget(self.resim_no_input, 1, 1)
        duzen_ozel.addWidget(QLabel("Ek Açıklama:"), 1, 2)
        duzen_ozel.addWidget(self.ek_aciklama_input, 1, 3)

        duzen_ozel.addWidget(QLabel("Kesilmis Mil Boyu:"), 0, 4)
        duzen_ozel.addWidget(self.kesilmis_mil_boyu, 0, 5)
        duzen_ozel.addWidget(QLabel("Mil Çapı:"), 1, 4)
        duzen_ozel.addWidget(self.mil_capi, 1, 5)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        # --- GRUP 3: Onay Kutuları (Checkbox) ---
        grup_onaylar = QGroupBox("Ek Donanımlar")
        duzen_onaylar = QHBoxLayout()

        self.cb_guduk = QCheckBox("GÜDÜK")
        self.cb_standart = QCheckBox("STANDART85")
        self.cb_govde_boyu_kullan = QCheckBox("Gövde Boyu Kullan")  # YENİ EKLENDİ

        duzen_onaylar.addWidget(self.cb_guduk)
        duzen_onaylar.addWidget(self.cb_standart)
        duzen_onaylar.addWidget(self.cb_govde_boyu_kullan)  # YENİ EKLENDİ


        grup_onaylar.setLayout(duzen_onaylar)
        form_ana_duzen.addWidget(grup_onaylar)


        scroll_area.setWidget(form_widget)
        ana_duzen.addWidget(scroll_area)
        self.form_alani_widgeti = scroll_area

        # ---------------- 3. BUTONLAR VE İŞLEMLER ----------------
        buton_duzeni = QHBoxLayout()

        self.btn_uret = QPushButton("LİSTEYE EKLE (KOD ÜRET)")
        self.btn_uret.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_uret.clicked.connect(self.uret_calistir)

        buton_duzeni.addWidget(self.btn_uret, stretch=2)
        ana_duzen.addLayout(buton_duzeni)

        # ---------------- 4. ARAMA VE TABLO ALANI ----------------
        arama_duzeni = QHBoxLayout()

        self.id_arama_kutusu = QLineEdit()
        self.id_arama_kutusu.setPlaceholderText("ID/Kod Ara...")

        self.arama_kutusu = QLineEdit()
        self.arama_kutusu.setPlaceholderText("Açıklamada Ara...")

        self.btn_ara = QPushButton("Ara")
        self.btn_ara.setStyleSheet(
            "background-color: #3498DB; color: white; font-weight: bold; padding: 5px; min-width: 80px;")
        self.btn_ara.clicked.connect(self.tabloda_arama_yap)

        self.btn_gelismis_ara = QPushButton("Gelişmiş Arama")
        self.btn_gelismis_ara.setStyleSheet("background-color: #E67E22; color: white; font-weight: bold; padding: 5px;")
        self.btn_gelismis_ara.clicked.connect(self.gelismis_arama_ac)

        self.btn_satir_sil = QPushButton("Seçili Satırı Sil")
        self.btn_satir_sil.setStyleSheet("background-color: #C0392B; color: white; font-weight: bold; padding: 5px;")
        self.btn_satir_sil.clicked.connect(self.secili_satiri_sil)

        self.btn_tam_ekran = QPushButton("🗖")
        self.btn_tam_ekran.setFixedSize(32, 32)
        self.btn_tam_ekran.setToolTip("Tabloyu Tam Ekran Yap / Küçült")
        self.btn_tam_ekran.setStyleSheet(
            "background-color: #8E44AD; color: white; font-weight: bold; font-size: 16px; border-radius: 4px;")
        self.btn_tam_ekran.clicked.connect(self.tam_ekran_modu_degistir)

        self.satir_sayaci = QLabel("Bulunan Kayıt: 0")
        self.satir_sayaci.setStyleSheet("font-weight: bold; color: #D35400; padding: 5px;")

        arama_duzeni.addWidget(QLabel("ID Ara:"))
        arama_duzeni.addWidget(self.id_arama_kutusu)
        arama_duzeni.addWidget(QLabel("Açıklama Ara:"))
        arama_duzeni.addWidget(self.arama_kutusu)
        arama_duzeni.addWidget(self.btn_ara)
        arama_duzeni.addWidget(self.btn_gelismis_ara)
        arama_duzeni.addWidget(self.btn_satir_sil)
        arama_duzeni.addWidget(self.btn_tam_ekran)
        arama_duzeni.addWidget(self.satir_sayaci)

        ana_duzen.addLayout(arama_duzeni)

        # ---------------- TABLO ----------------
        self.tablo = QTableWidget(0, 8)
        self.tablo_basliklari = ["#KOD", "AÇIKLAMA", "GRUP", 'AÇIKLAMA2', "RESİM NO", 'AÇIKLAMA SIRA NO',
                                 'AÇIKLAMA SIRA KOD', "KAYIT TARİHİ"]
        self.tablo.setHorizontalHeaderLabels(self.tablo_basliklari)
        # Sütunların fare ile genişletilip daraltılmasına izin verir (Çift tıklayınca otomatik sığdırır)
        self.tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.tablo.horizontalHeader().setStretchLastSection(True)  # Sadece son sütun kalan boşluğu doldursun
        self.tablo.setWordWrap(False)  # Uzun yazıların alt satıra geçip satırı kalınlaştırmasını engeller

        # Başlığa sağ tık menüsü (Sütun sığdırma için)
        self.tablo.horizontalHeader().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tablo.horizontalHeader().customContextMenuRequested.connect(self.baslik_sag_tik_menusu)
        self.tablo.setAlternatingRowColors(True)
        # Filtreleme / Sıralama
        self.tablo.setSortingEnabled(True)

        # Sağ tık ve kopyalama özellikleri
        self.tablo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tablo.customContextMenuRequested.connect(self.tablo_sag_tik)
        self.tablo.keyPressEvent = self.tablo_kopyalama_dinleyicisi
        ana_duzen.addWidget(self.tablo, stretch=1)

        # ---------------- EXCEL'E AKTAR BUTONU ----------------
        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR ( MIL)")
        self.aktar_butonu.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.aktar_butonu.clicked.connect(self.excele_aktar)
        ana_duzen.addWidget(self.aktar_butonu)

        # --- BÜTÜN COMBOBOX'LARI MANUEL YAZIMA AÇMA MOTORU ---
        # --- BÜTÜN COMBOBOX'LARI MANUEL YAZIMA VE ANINDA SİLMEYE AÇMA MOTORU ---
        for kutu in self.findChildren(QComboBox):
            kutu.setEditable(True)
            line_edit = kutu.lineEdit()

            # Kutunun içine ipucu metni ekler
            line_edit.setPlaceholderText("Seç veya Yaz...")

            # Tıklandığı an içindeki bütün metni blok halinde seçer!
            # Böylece adam bir harfe bastığı an eski metin anında yok olur.
            line_edit.mousePressEvent = lambda event, le=line_edit: le.selectAll()

            # Otomatik tamamlama ayarları (Özellikle Ana Motor ve uzun listeler için çok faydalı)
            kutu.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            completer = kutu.completer()
            if completer:
                completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                completer.setFilterMode(Qt.MatchFlag.MatchContains)

        # SAAT İÇİN ZAMANLAYICI
        self.zamanlayici = QTimer(self)
        self.zamanlayici.timeout.connect(self.tarih_mevcut_gosterim)
        self.zamanlayici.start(1000)

    # =========================================================================
    # FONKSİYONLAR
    # =========================================================================

    def tarih_mevcut_gosterim(self):
        suan = QDateTime.currentDateTime()
        self.saat_etiketi.setText(suan.toString("dd.MM.yyyy HH:mm:ss"))

    def formu_temizle(self):
        for nesne in self.findChildren(QComboBox):
            nesne.setCurrentIndex(0)
        for nesne in self.findChildren(QLineEdit):
            nesne.clear()  # Artık arama kutularını da siliyor!
        for nesne in self.findChildren(QCheckBox):
            nesne.setChecked(False)
        self.tablo.setRowCount(0)
        self.satir_sayaci.setText("Bulunan Kayıt: 0")

    def tam_ekran_modu_degistir(self):
        su_an_gorunur_mu = self.form_alani_widgeti.isVisible()
        self.form_alani_widgeti.setVisible(not su_an_gorunur_mu)
        if su_an_gorunur_mu:
            self.btn_tam_ekran.setText("Geri Getir")
            self.btn_tam_ekran.setStyleSheet(
                "background-color: #27AE60; color: white; font-weight: bold; font-size: 10px; border-radius: 4px;")
        else:
            self.btn_tam_ekran.setText("🗖")
            self.btn_tam_ekran.setStyleSheet(
                "background-color: #8E44AD; color: white; font-weight: bold; font-size: 16px; border-radius: 4px;")

    def secili_satiri_sil(self):
        secili_satirlar = self.tablo.selectionModel().selectedRows()
        if not secili_satirlar:
            QMessageBox.warning(self, "Uyarı", "Lütfen silmek için tablodan bir satır seçin!")
            return
        for satir in sorted(secili_satirlar, reverse=True):
            self.tablo.removeRow(satir.row())

    def genel_max_sira_no_bul(self):
        kullanilan_nolar = set()

        # 1. EXCEL'DEKİ KULLANILMIŞ NUMARALARI TOPLA
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = "MIL"
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.strip().upper() == "MIL":
                    hedef_sayfa = sayfa
                    break

            # DİREKT HEADER=0 İLE OKU VE DÖNGÜSÜZ TARAMAYA GEÇ
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna("")
            for idx, row in df.iterrows():
                for val in row.values:
                    val_str = str(val).strip().upper()
                    if "ML-" in val_str:
                        parcalar = val_str.split("-")
                        for parca in parcalar:
                            if parca.isdigit() and len(parca) == 4:
                                kullanilan_nolar.add(int(parca))
        except Exception as e:
            pass

        # 2. ARAYÜZ TABLOSUNDAKİ (HENÜZ KAYDEDİLMEMİŞ) NUMARALARI TOPLA
        for row_idx in range(self.tablo.rowCount()):
            try:
                tablo_id = self.tablo.item(row_idx, 0).text().strip().upper()
                if "ML-" in tablo_id:
                    parcalar = tablo_id.split("-")
                    for parca in parcalar:
                        if parca.isdigit() and len(parca) == 4:
                            kullanilan_nolar.add(int(parca))
            except:
                pass

        # 3. İLK BOŞLUĞU BUL
        for i in range(1, 10000):
            if i not in kullanilan_nolar:
                return i

        return len(kullanilan_nolar) + 1

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip().upper())

        if not aranan_id and not aranan_aciklama:
            return

        guncel_yol = ayar_excel_yolunu_getir()
        self.tablo.setRowCount(0)


        # ---------------- AKILLI SAYFA BULUCU ----------------
        gercek_sayfa_adi = "MIL"
        try:
            excel_dosyasi = pd.ExcelFile(guncel_yol)
            for sayfa in excel_dosyasi.sheet_names:
                temiz_sayfa_adi = sayfa.strip().upper()
                if temiz_sayfa_adi == "MIL" or temiz_sayfa_adi == "MİL":
                    gercek_sayfa_adi = sayfa
                    break
        except:
            pass
        try:
            # YENİ VE TEMİZ KISIM (Döngüler Silindi)
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=gercek_sayfa_adi, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]

            df_veri = df.copy()  # Artık başlık satırı atlamaya gerek yok
            bulunan_kayit_sayisi = 0
            self.tablo.setSortingEnabled(False)

            for index, row in df_veri.iterrows():

                ilk_hucre = str(row.iloc[0]).strip()
                if not ilk_hucre or ilk_hucre == 'nan' or "KOD" in ilk_hucre.upper():
                    continue

                # Satırdaki tüm metinleri birleştirip devasa bir arama havuzu oluşturuyoruz
                tum_satir = " ".join([str(val).strip() for val in row.values if pd.notna(val)])
                tam_havuz = turkce_karakter_temizleme(tum_satir).upper()

                id_eslesti = True
                aciklama_eslesti = True

                # ID kutusunda arama (063, RM-, ML- vb. esnek arama)
                if aranan_id and aranan_id not in tam_havuz:
                    id_eslesti = False

                # Açıklama kutusunda kelime bazlı esnek arama
                if aranan_aciklama:
                    aranan_kelimeler = aranan_aciklama.split()
                    if not all(kelime in tam_havuz for kelime in aranan_kelimeler):
                        aciklama_eslesti = False

                if id_eslesti and aciklama_eslesti:
                    mevcut_satir = self.tablo.rowCount()
                    self.tablo.insertRow(mevcut_satir)

                    # İlk 8 sütunu tablonun 0-7 indislerine tam oturtuyoruz
                    for col_idx in range(min(8, len(row))):
                        veri = row.iloc[col_idx]
                        if pd.notna(veri) and str(veri).strip() != 'nan':
                            self.tablo.setItem(mevcut_satir, col_idx, QTableWidgetItem(str(veri).strip()))

                    bulunan_kayit_sayisi += 1

            self.tablo.setSortingEnabled(True)
            self.satir_sayaci.setText(f"Bulunan Kayıt: {bulunan_kayit_sayisi}")

            # YENİ: BULUNAMADI UYARISI
            if bulunan_kayit_sayisi == 0:
                QMessageBox.information(self, "Bulunamadı", "Aradığınız kriterlerde kayıt bulunamadı.")

        except ValueError:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Excel dosyasında sayfa bulunamadı!")
        except Exception as e:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Arama sırasında hata:\n{e}")


    def gelismis_arama_ac(self):
        try:
            excel_yolu = ayar_excel_yolunu_getir()
            dialog = MilGelismisArama(self, excel_yolu)
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Gelişmiş arama açılamadı:\n{e}")

    def excele_aktar(self):
        gercek_sayfa_adi = "MIL"
        guncel_yol = ayar_excel_yolunu_getir()
        excel_yolu = os.path.abspath(guncel_yol)

        if self.tablo.rowCount() == 0:
            QMessageBox.warning(self, "Uyarı", "Tabloda aktarılacak veri yok!")
            return


        # 1. KOPYA KONTROLÜ
        try:
            sayfa_df = pd.read_excel(excel_yolu, sheet_name=gercek_sayfa_adi, header=0, dtype=str).fillna("")
            sayfa_df.columns = [str(c).strip().upper() for c in sayfa_df.columns]

            id_kolonu = None
            for col in sayfa_df.columns:
                if "KOD" in col and "AÇIKLAMA" not in col:
                    id_kolonu = col
                    break

            exceldeki_idler = sayfa_df[id_kolonu].astype(str).str.strip().tolist() if id_kolonu else []
            cakisan_satirlar = []
            for satir in range(self.tablo.rowCount()):
                item = self.tablo.item(satir, 0)
                if item and item.text().strip() in exceldeki_idler:
                    cakisan_satirlar.append(satir)
        # ... (SADECE EVET VE İPTAL SEÇENEĞİ diyerek devam eden kısım aynı kalacak)

            # SADECE EVET VE İPTAL SEÇENEĞİ (Notlarına uygun)
            if cakisan_satirlar:
                cevap_kopya = QMessageBox.question(
                    self, "KOPYA MİL UYARISI!",
                    f"Tablodaki {len(cakisan_satirlar)} adet kayıt Excel'de zaten mevcut!\n\n"
                    f"• EVET: Kopyaları tablodan kaldır ve sadece yenileri kaydet.\n"
                    f"• İPTAL: İşlemi tamamen durdur.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel
                )
                if cevap_kopya == QMessageBox.StandardButton.Cancel:
                    return
                elif cevap_kopya == QMessageBox.StandardButton.Yes:
                    for r in sorted(cakisan_satirlar, reverse=True): self.tablo.removeRow(r)
                    if self.tablo.rowCount() == 0:
                        QMessageBox.information(self, "Bilgi",
                                                "Kopyalar silinince tabloda kaydedilecek yeni veri kalmadı.")
                        return
        except Exception as e:
            pass

            # 2. KAYIT ONAYI
        excel_aktarim_cevap = QMessageBox.question(self, "KAYIT ONAYI",
                                                   f"Tablodaki verileri Excel'e ({gercek_sayfa_adi}) GÖNDERMEK istiyor musunuz?",
                                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if excel_aktarim_cevap == QMessageBox.StandardButton.No:
            return

        # 3. GÖLGE KOPYA (SHADOW COPY) VE YAZMA SİSTEMİ
        try:
            temp_klasor = tempfile.gettempdir()
            gecici_excel = os.path.join(temp_klasor, "gecici_aemot_mil.xlsx")
            shutil.copy2(excel_yolu, gecici_excel)
            os.chmod(gecici_excel, stat.S_IWRITE)

            app = xw.App(visible=False)
            app.display_alerts = False

            wb = app.books.open(gecici_excel)
            sheet = wb.sheets[gercek_sayfa_adi]


            # 1. Satır H ve I Sütunlarına başlıkları yaz (A1 Standartı)
            sheet.range('H1').value = "KAYIT TARİHİ"
            sheet.range('I1').value = "KAYDI EKLEYEN KİŞİ"
            try:
                sheet.range('A1:I1').api.Font.Bold = True
            except:
                pass

            son_satir = sheet.range('A1048576').end('up').row
            if son_satir < 1: son_satir = 1

            aktarilacak_veriler = []
            for r in range(self.tablo.rowCount()):
                satir_verisi = []
                # 8 Sütunluk veriyi çek
                for c in range(8):
                    item = self.tablo.item(r, c)
                    satir_verisi.append(item.text() if item else "")

                # SATIRIN SONUNA KULLANICIYI ZIMBALA
                satir_verisi.append(oturum.tam_isim())
                aktarilacak_veriler.append(satir_verisi)

            if aktarilacak_veriler:
                baslama_satiri = son_satir + 1
                bitis_satiri = baslama_satiri + len(aktarilacak_veriler) - 1

                # Hedef alanı I sütununa kadar genişlettik!
                hedef_alan = sheet.range(f'A{baslama_satiri}:I{bitis_satiri}')
                hedef_alan.number_format = '@'
                hedef_alan.value = aktarilacak_veriler

            wb.save(gecici_excel)
            wb.close()
            app.quit()

            # Ağdaki dosyanın kilidini kır ve üstüne yaz
            try:
                if os.path.exists(excel_yolu): os.chmod(excel_yolu, stat.S_IWRITE)
                shutil.copy2(gecici_excel, excel_yolu)

                QMessageBox.information(self, "Başarılı",
                                        "Kayıtlar Excel'in 9. satırından itibaren kusursuzca eklendi!")
                self.tablo.setRowCount(0)

            except PermissionError:
                QMessageBox.critical(self, "DOSYA KİLİTLİ!",
                                     "Veriler hazırlandı ancak Excel dosyası şu an AÇIK olduğu için yazılamadı. Lütfen kapatıp tekrar deneyin.")
            except Exception as e:
                QMessageBox.critical(self, "Aktarım Hatası", f"Dosya yerine konurken hata oluştu:\n{e}")

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            QMessageBox.critical(self, "Lokal İşlem Hatası", f"Veri işlenirken bir hata oluştu:\n{e}")

    def dna_bazli_sira_no_bul(self, sayfa_adi, aranan_grup, taslak_metin, resim_no_cek=False):

        def kelime_dizisi_yap(metin):
            temiz_metin = turkce_karakter_temizleme(str(metin)).upper()  # DÜZELTME: TR karakter normalize edilmeden karşılaştırılıyordu

            # --- YENİ EKLENEN KURAL: REVİZYON NUMARALARINI DNA'DAN SİL! ---
            # Böylece R:01 ile R:05 olan parçalar sırf revizyonları farklı diye ayrı sıra no almaz.
            import re
            temiz_metin = re.sub(r'R\s*:\s*\d+', '', temiz_metin)

            for isaret in ["-", ":", "/", "\\", "_", ",", ".", "MM"]:
                temiz_metin = temiz_metin.replace(isaret, " ")
            return set(temiz_metin.split())

        programin_dizisi = kelime_dizisi_yap(taslak_metin)

        # --- 1. ARAYÜZ TABLOSU ---
        for row_idx in range(self.tablo.rowCount()):
            tablo_grup = str(self.tablo.item(row_idx, 2).text()).strip().upper()

            if kelime_dizisi_yap(tablo_grup) == kelime_dizisi_yap(aranan_grup):
                tab_aciklama = str(self.tablo.item(row_idx, 1).text())
                tab_aciklama2 = str(self.tablo.item(row_idx, 3).text())
                tablo_dizisi = kelime_dizisi_yap(f"{tab_aciklama} {tab_aciklama2}")

                if programin_dizisi == tablo_dizisi:
                    sira_no = int(self.tablo.item(row_idx, 5).text().strip())
                    if resim_no_cek:
                        tab_resim = str(self.tablo.item(row_idx, 4).text()).strip()
                        return (sira_no, tab_resim)
                    return sira_no

        # --- 2. EXCEL GEÇMİŞİ ---
        # --- 2. EXCEL GEÇMİŞİ ---
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = sayfa_adi
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.upper().strip() == "MIL" or sayfa.upper().strip() == "MİL":
                    hedef_sayfa = sayfa
                    break

            # YENİ VE TEMİZ KISIM (Döngüler çöpe)
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna(
                "")
            df.columns = [str(c).strip().upper() for c in df.columns]

            for index, row in df.iterrows():
                excel_grup = str(row.get('GRUP', ''))
                # ... (Geri kalan if blokları aynı kalacak)

                if kelime_dizisi_yap(excel_grup) == kelime_dizisi_yap(aranan_grup):
                    excel_aciklama = str(row.get('AÇIKLAMA', ''))
                    excel_aciklama2 = str(row.get('AÇIKLAMA2', row.get('EK AÇIKLAMA', '')))
                    excel_dizisi = kelime_dizisi_yap(f"{excel_aciklama} {excel_aciklama2}")

                    if programin_dizisi == excel_dizisi:
                        sira_kolonu = 'AÇIKLAMA SIRA NO'
                        if sira_kolonu not in df.columns:
                            for col in df.columns:
                                if "SIRA" in col or "NO" in col and "RESİM" not in col:
                                    sira_kolonu = col
                                    break

                        sira_no = row.get(sira_kolonu, None)
                        try:
                            sira_no = int(float(str(sira_no).replace("'", "").strip()))
                            if resim_no_cek:
                                resim = str(row.get('RESİM NO', '')).strip()
                                return (sira_no, resim)
                            return sira_no
                        except:
                            pass
        except Exception as e:
            print(f"DNA Okuma Hatası: {e}")

        # Eşleşme yoksa (Yepyeni bir özellikse)
        return (None, "") if resim_no_cek else None

    def uret_calistir(self):
        # --- KULLANICININ MANUEL YAZDIĞI HER ŞEYİ AÇILIR LİSTE HAFIZASINA EKLE ---
        for kutu in self.findChildren(QComboBox):
            yeni_metin = kutu.currentText().strip().upper()
            if yeni_metin and kutu.findText(yeni_metin) == -1:
                kutu.addItem(yeni_metin)

        # 1. ARAYÜZDEN VERİLERİ ÇEK
        versiyon = self.versiyon_kutusu.currentText()
        grup = self.grup_kutusu.currentText()
        kutup = self.kutup_kutusu.currentText()
        govde_boyu = self.govde_boyu_kutusu.currentText()
        kesilmis_islenmis = self.kesilmis_islenmis.currentText()
        revizyon = self.revizyon_kutusu.currentText()
        resim_no_girdi = self.resim_no_input.text().strip()
        mil_boyu = self.kesilmis_mil_boyu.text().strip().upper()
        mil_capi = self.mil_capi.text().strip().upper()
        ham_ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
        ek_aciklama = ham_ek_aciklama.replace(" MIL", "").replace(" MİL", "").strip()

        # --- BOŞ ALAN UYARISI ---
        bos_alanlar = []
        if not grup: bos_alanlar.append("Grup")
        if not govde_boyu: bos_alanlar.append("Gövde Boyu")
        if not kesilmis_islenmis: bos_alanlar.append("Kesilmiş/İşlenmiş")

        if bos_alanlar:
            cevap = QMessageBox.question(self, "Eksik Bilgi",
                                         f"Şu alanlar boş bırakıldı:\n{', '.join(bos_alanlar)}\n\nYine de üretip tabloya eklemek istiyor musunuz?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        if not resim_no_girdi:
            resim_no_girdi = "0000"

        guduk = self.cb_guduk.isChecked()
        standart = self.cb_standart.isChecked()

        # ---------------- 2. KUSURSUZ TASLAK OLUŞTURMA MANTIĞI ----------------
        # ---------------- 2. KUSURSUZ TASLAK OLUŞTURMA MANTIĞI ----------------
        grup_kutup = f"{grup}-{kutup}" if kutup else grup

        # KURAL: Resim No sadece ISLENMIS seçilmişse geçerlidir! KESILMIS ise "0000" kalır.
        if kesilmis_islenmis == "KESILMIS":
            resim_no_girdi = "0000"

        if not resim_no_girdi:
            resim_no_girdi = "0000"

        # Gövde boyu kullanım durumu (Checkbox kontrolü)
        kullanilacak_boy = govde_boyu
        if getattr(self, 'cb_govde_boyu_kullan', None) and self.cb_govde_boyu_kullan.isChecked():
            kullanilacak_boy = govde_boyu
        else:
            kullanilacak_boy = f"{mil_boyu}MM" if mil_boyu else govde_boyu

        guduk = self.cb_guduk.isChecked()
        standart = getattr(self, 'cb_standart', None) and self.cb_standart.isChecked()

        # KURAL 1: Standart Çap Hesaplama
        hesaplanan_cap = ""
        if kesilmis_islenmis == "KESILMIS" and not guduk:
            if standart:
                try:
                    cap_tablosu = cap_oku()
                    ozel_kod = f"{grup}-{kutup}" if kutup else grup
                    if ozel_kod in cap_tablosu:
                        hesaplanan_cap = cap_tablosu[ozel_kod]
                    elif grup in cap_tablosu:
                        hesaplanan_cap = cap_tablosu[grup]
                    else:
                        hesaplanan_cap = "???"
                except:
                    hesaplanan_cap = "???"
            else:
                hesaplanan_cap = mil_capi

        # --- Temel Açıklama Çatısını Kuruyoruz ---
        taslak_aciklama = ""
        saf_dna_metni = ""  # Açıklama 2 ve Sıra No bulmak için (Resim Nosuz ve Revizyonsuz)

        if guduk:
            # KURAL 4: Güdük İstisnası (Her şeyi ezer)
            kutup_eki = f"-{kutup}" if kutup else ""
            saf_dna_metni = f"{grup} {govde_boyu}{kutup_eki} GUDUK {kesilmis_islenmis} MIL"

            # GÜDÜK KURALI: Resim no kesinlikle açıklamada YOK!
            taslak_aciklama = saf_dna_metni

        elif kesilmis_islenmis == "ISLENMIS":
            # KURAL 3: İşlenmiş Durumu
            if versiyon == "ESKİ_GÖVDE":
                saf_dna_metni = f"{grup_kutup} {govde_boyu} ISLENMIS MIL" if govde_boyu else f"{grup_kutup} ISLENMIS MIL"
                if resim_no_girdi != "0000":
                    # Eski gövdede resim no varsa gövde boyu ezilir, kutuptan sonra resim no gelir
                    taslak_aciklama = f"{grup_kutup} {resim_no_girdi} ISLENMIS MIL"
                else:
                    taslak_aciklama = saf_dna_metni
            else:
                # PREMIUM İŞLENMİŞ: Gövde Boyu EZİLİR. (Örn: 071 ISLENMIS MIL)
                saf_dna_metni = f"{grup_kutup} ISLENMIS MIL"
                taslak_aciklama = saf_dna_metni

        else:
            # Standart Kesilmiş Durumu
            cap_metni = f" - Ø:{hesaplanan_cap}MM" if hesaplanan_cap else ""
            saf_dna_metni = f"{grup_kutup} {kullanilacak_boy}{cap_metni} KESILMIS MIL"

            if versiyon == "ESKİ_GÖVDE" and resim_no_girdi != "0000":
                # Eski gövdede kutuptan sonra resim no gelir
                taslak_aciklama = f"{grup_kutup} {resim_no_girdi} {kullanilacak_boy}{cap_metni} KESILMIS MIL"
            else:
                taslak_aciklama = saf_dna_metni

        # Ek Açıklama Varsa Sona Ekle
        if ek_aciklama:
            saf_dna_metni += f" {ek_aciklama}"
            taslak_aciklama += f" {ek_aciklama}"

        # KURAL 5: Premium Revizyon Eklentisi
        if versiyon == "PREMIUM":
            taslak_aciklama += f" - R:{revizyon}"

        taslak_aciklama = " ".join(taslak_aciklama.split())
        saf_dna_metni = " ".join(saf_dna_metni.split())

        # ---------------- 3. DNA TESTİ İLE SIRA VE RESİM NO BULMA ----------------
        tam_metin_havuzu = saf_dna_metni

        dna_sira, dna_resim = self.dna_bazli_sira_no_bul(sayfa_adi="MIL", aranan_grup=grup,
                                                         taslak_metin=tam_metin_havuzu, resim_no_cek=True)

        resim_no = resim_no_girdi
        if (not resim_no or resim_no == "0000") and dna_resim:
            resim_no = str(dna_resim).strip()

        if resim_no.isdigit():
            resim_no = resim_no.zfill(4)
        if not resim_no or resim_no == "nan":
            resim_no = "0000"

        from arayuz.mantik.sira_no_dialog import sira_no_secim_yap
        secilen_sira = sira_no_secim_yap(self, dna_sira, taslak_aciklama, sayfa_adi="MIL", id_kodu="ML")
        if secilen_sira is None:
            return
        elif secilen_sira == -1:
            sira_no_int = self.genel_max_sira_no_bul()  # DÜZELTME: fonksiyon zaten ilk boş sıra noyu döndürüyor, +1 eklemek onu dolu bir numaraya kaydırıyordu
        else:
            sira_no_int = secilen_sira

        sira_no_str = str(sira_no_int)
        aciklama_kod_str = sira_no_str.zfill(4)

        # ---------------- 4. NİHAİ ID VE AÇIKLAMAYI OLUŞTURMA ----------------
        id_revizyon = "00" if versiyon == "ESKİ_GÖVDE" else revizyon

        if kesilmis_islenmis == "KESILMIS" or resim_no == "0000":
            id_resim_no = "00000"
        else:
            try:
                if int(resim_no) < 10000:
                    id_resim_no = f"0{resim_no}"
                else:
                    id_resim_no = resim_no
            except:
                id_resim_no = resim_no

        uretilen_id = f"YMI{grup}-ML-{aciklama_kod_str}-{id_revizyon}-{id_resim_no}"

        # --- KOPYA KALKANI ---
        for r in range(self.tablo.rowCount()):
            if self.tablo.item(r, 0) and self.tablo.item(r, 0).text().strip() == uretilen_id:
                QMessageBox.warning(self, "KOPYA ENGELİ",
                                    f"Bu Mil ZATEN LİSTEDE var!\nID: {uretilen_id}\n\nAynı mili ard arda ekleyemezsiniz.")
                return

        # NİHAİ AÇIKLAMA OLUŞTURMA
        uretilen_aciklama = taslak_aciklama

        # Premium için resim no revizyondan sonra en sona eklenir
        # FAKAT GÜDÜK SEÇİLİYSE EKLENMEZ!
        if versiyon == "PREMIUM" and resim_no != "0000" and not guduk:
            uretilen_aciklama += f" - {resim_no}"

        uretilen_aciklama = " ".join(uretilen_aciklama.split())

        # --- AÇIKLAMA 2'Yİ OTOMATİK ÜRETMESİ ---
        # Açıklama 2'de kesinlikle resim no ve revizyon yer almayacak.
        # Bu yüzden içinde resim no ve revizyon barındırmayan saf_dna_metni'ni kullanıyoruz.
        aciklama2_oto = saf_dna_metni

        if aciklama2_oto.startswith(grup):
            aciklama2_oto = aciklama2_oto[len(grup):].strip()

        # ---------------- 5. TABLOYA YAZDIRMA ----------------
        bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

        # DİKKAT: Tabloya manuel yazılan ek_aciklama DEĞİL, sistemin oluşturduğu aciklama2_oto gidiyor!
        yeni_kayit = [
            uretilen_id, uretilen_aciklama, grup, aciklama2_oto, resim_no, sira_no_str, aciklama_kod_str, bugunun_tarihi
        ]

        self.tablo.setSortingEnabled(False)
        mevcut_satir = self.tablo.rowCount()
        self.tablo.insertRow(mevcut_satir)
        for sutun, veri in enumerate(yeni_kayit):
            hucre = QTableWidgetItem(str(veri))
            self.tablo.setItem(mevcut_satir, sutun, hucre)
        self.tablo.setSortingEnabled(True)

    def tablo_sag_tik(self, pos):
        from PyQt6.QtWidgets import QMenu
        menu = QMenu()
        kopyala_aksiyon = menu.addAction("Seçili Alanı Kopyala (Ctrl+C)")
        secim = menu.exec(self.tablo.viewport().mapToGlobal(pos))
        if secim == kopyala_aksiyon: self.kopyala_islev()

    def tablo_kopyalama_dinleyicisi(self, event):
        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            self.kopyala_islev()
        else:
            super(QTableWidget, self.tablo).keyPressEvent(event)

    def kopyala_islev(self):
        from PyQt6.QtWidgets import QApplication
        secili_hucreler = self.tablo.selectedIndexes()
        if not secili_hucreler: return
        secili_hucreler.sort(key=lambda index: (index.row(), index.column()))
        kopyalanacak_metin = ""
        onceki_satir = secili_hucreler[0].row()
        for hucre in secili_hucreler:
            mevcut_satir = hucre.row()
            metin = hucre.data(Qt.ItemDataRole.DisplayRole)
            if mevcut_satir != onceki_satir:
                kopyalanacak_metin += "\n"
                onceki_satir = mevcut_satir
            else:
                if hucre != secili_hucreler[0]: kopyalanacak_metin += "\t"
            kopyalanacak_metin += str(metin) if metin else ""
        QApplication.clipboard().setText(kopyalanacak_metin)

    def baslik_sag_tik_menusu(self, pos):

        menu = QMenu()

        aksiyon_sutun_sigdir = menu.addAction("Bu Sütunu İçeriğe Sığdır")
        aksiyon_tum_sutunlar = menu.addAction("Tüm Sütunları İçeriğe Sığdır")

        # Hangi sütun başlığına tıklandığını bul
        sutun_index = self.tablo.horizontalHeader().logicalIndexAt(pos)

        secim = menu.exec(self.tablo.horizontalHeader().mapToGlobal(pos))

        if secim == aksiyon_sutun_sigdir:
            self.tablo.resizeColumnToContents(sutun_index)  # Sadece tıklanan sütunu sığdır
        elif secim == aksiyon_tum_sutunlar:
            self.tablo.resizeColumnsToContents()  # Excel'deki gibi tüm tabloyu sığdır


class MilGelismisArama(QDialog):
    def __init__(self, parent, excel_yolu):
        super().__init__(parent)
        self.setWindowTitle("Mil Gelişmiş Arama")
        self.setMinimumWidth(700)
        self.parent = parent
        self.excel_yolu = excel_yolu

        self.layout = QVBoxLayout(self)

        # --- 1. EXCEL VERİSİNİ OKU ---
        gercek_sayfa = "MIL"
        try:
            self.df = pd.read_excel(self.excel_yolu, sheet_name=gercek_sayfa, header=0, dtype=str).fillna("")
            self.df.columns = [str(c).strip().upper() for c in self.df.columns]
            self.temiz_kolonlar = {str(col): col for col in self.df.columns}
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"MIL sayfası bulunamadı veya okunamadı!\n{e}")
            self.df = pd.DataFrame()

        # --- 2. FORM ALANI ---
        grup_form = QGroupBox("MİL Parametreleriyle Filtrele")
        form_layout = QGridLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["", "ESKİ_GÖVDE", "PREMIUM"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["", "063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315",
             "355", "400", "450", "500", "630", "PILOTCAR"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(["", "2", "4", "6", "8", "10", "12", "2/4", "4/2", "4/6"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(
            ["", "S", "M", "L", "S-M", "S-L", "M-L", "S-M-L", "260PK", "280PK", "300PK", "310PK", "340Pk"
                , "360PK", "400PK", "420PK", "450PK", "470PK", "475PK", "520PK", "525PK", "550PK", "585PK"
                , "650PK", "660PK", "685PK", "750PK", "920PK", "M-725PK"])

        self.kesilmis_islenmis = QComboBox()
        self.kesilmis_islenmis.addItems(["", "KESILMIS", "ISLENMIS"])

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["", "01", "02", "03", "04", "05"])

        self.resim_no_input = QLineEdit()
        self.resim_no_input.setPlaceholderText("Örn: 3130")

        self.kesilmis_mil_boyu = QLineEdit()
        self.kesilmis_mil_boyu.setPlaceholderText("Örn: 850")

        self.mil_capi = QLineEdit()
        self.mil_capi.setPlaceholderText("Örn: 45")

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Örn: BAKIR")

        self.sira_no_kutusu = QLineEdit()
        self.sira_no_kutusu.setPlaceholderText("Örn: 1")

        form_layout.addWidget(QLabel("Versiyon:"), 0, 0)
        form_layout.addWidget(self.versiyon_kutusu, 0, 1)
        form_layout.addWidget(QLabel("Grup (M):"), 0, 2)
        form_layout.addWidget(self.grup_kutusu, 0, 3)

        form_layout.addWidget(QLabel("Kutup:"), 1, 0)
        form_layout.addWidget(self.kutup_kutusu, 1, 1)
        form_layout.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        form_layout.addWidget(self.govde_boyu_kutusu, 1, 3)

        form_layout.addWidget(QLabel("Kesilmiş/İşlenmiş:"), 2, 0)
        form_layout.addWidget(self.kesilmis_islenmis, 2, 1)
        form_layout.addWidget(QLabel("Revizyon No:"), 2, 2)
        form_layout.addWidget(self.revizyon_kutusu, 2, 3)

        form_layout.addWidget(QLabel("Resim No:"), 3, 0)
        form_layout.addWidget(self.resim_no_input, 3, 1)
        form_layout.addWidget(QLabel("Ek Açıklama:"), 3, 2)
        form_layout.addWidget(self.ek_aciklama_input, 3, 3)

        form_layout.addWidget(QLabel("Mil Boyu:"), 4, 0)
        form_layout.addWidget(self.kesilmis_mil_boyu, 4, 1)
        form_layout.addWidget(QLabel("Mil Çapı:"), 4, 2)
        form_layout.addWidget(self.mil_capi, 4, 3)

        form_layout.addWidget(QLabel("Sıra No (Kod):"), 5, 0)
        form_layout.addWidget(self.sira_no_kutusu, 5, 1)

        self.cb_guduk = QCheckBox("GÜDÜK")
        self.cb_standart = QCheckBox("STANDART85")
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.cb_guduk)
        checkbox_layout.addWidget(self.cb_standart)
        form_layout.addLayout(checkbox_layout, 6, 0, 1, 4)

        grup_form.setLayout(form_layout)
        self.layout.addWidget(grup_form)

        # --- 3. TARİH ARALIĞI ---
        grup_tarih = QGroupBox("Kayıt Tarihi Aralığı")
        duzen_tarih = QHBoxLayout()

        self.cb_tarih_kullan = QCheckBox("Tarihe Göre Ara")
        self.bas_tarih = QDateEdit()
        self.bas_tarih.setCalendarPopup(True)
        self.bas_tarih.setDate(QDate(2025, 1, 1))
        self.bas_tarih.setDisplayFormat("dd.MM.yyyy")
        self.bas_tarih.setEnabled(False)

        self.bit_tarih = QDateEdit()
        self.bit_tarih.setCalendarPopup(True)
        self.bit_tarih.setDate(QDate.currentDate())
        self.bit_tarih.setDisplayFormat("dd.MM.yyyy")
        self.bit_tarih.setEnabled(False)

        self.cb_tarih_kullan.toggled.connect(self.bas_tarih.setEnabled)
        self.cb_tarih_kullan.toggled.connect(self.bit_tarih.setEnabled)

        duzen_tarih.addWidget(self.cb_tarih_kullan)
        duzen_tarih.addWidget(QLabel("Başlangıç:"))
        duzen_tarih.addWidget(self.bas_tarih)
        duzen_tarih.addWidget(QLabel("Bitiş:"))
        duzen_tarih.addWidget(self.bit_tarih)
        grup_tarih.setLayout(duzen_tarih)
        self.layout.addWidget(grup_tarih)

        self.btn_ara = QPushButton("Filtrele ve Tabloya Getir")
        self.btn_ara.setStyleSheet("background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.btn_ara.clicked.connect(self.arama_yap)
        self.layout.addWidget(self.btn_ara)

        for kutu in self.findChildren(QComboBox):
            kutu.setEditable(True)
            kutu.lineEdit().mousePressEvent = lambda event, le=kutu.lineEdit(): le.selectAll()

    def arama_yap(self):
        try:
            self.parent.tablo.setSortingEnabled(False)
            df_filtrelenmis = self.df.copy()

            if df_filtrelenmis.empty:
                self.parent.tablo.setSortingEnabled(True)
                self.accept()
                return

            tum_metin = pd.Series([""] * len(df_filtrelenmis), index=df_filtrelenmis.index)
            for col_idx in range(min(10, len(df_filtrelenmis.columns))):
                tum_metin += df_filtrelenmis.iloc[:, col_idx].astype(str).str.upper() + " "

            tum_metin = tum_metin.str.replace("İ", "I").str.replace("Ş", "S").str.replace("Ğ", "G").str.replace("Ü", "U").str.replace("Ö", "O").str.replace("Ç", "C")
            metin_sifir_bosluk = tum_metin.str.replace(" ", "").str.replace("-", "").str.replace("/", "")

            versiyon = self.versiyon_kutusu.currentText().strip().upper()
            grup = self.grup_kutusu.currentText().strip().upper()
            kutup = self.kutup_kutusu.currentText().strip().upper()
            govde_boyu = self.govde_boyu_kutusu.currentText().strip().upper()
            kes_isl = self.kesilmis_islenmis.currentText().strip().upper()
            revizyon = self.revizyon_kutusu.currentText().strip().upper()
            resim_no = self.resim_no_input.text().strip().upper()
            mil_boyu = self.kesilmis_mil_boyu.text().strip().upper()
            mil_cap = self.mil_capi.text().strip().upper()
            ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
            sira_no = self.sira_no_kutusu.text().strip()
            guduk_secili = self.cb_guduk.isChecked()

            if grup:
                if grup == "PILOTCAR":
                    df_filtrelenmis = df_filtrelenmis[
                        (df_filtrelenmis.iloc[:, 2].astype(str).str.strip().str.upper().isin(["PILOTCAR", "PIL"])) |
                        (metin_sifir_bosluk.str.contains("PILOTCAR", regex=False))
                    ]
                else:
                    df_filtrelenmis = df_filtrelenmis[df_filtrelenmis.iloc[:, 2].astype(str).str.strip().str.upper() == grup]

            if resim_no:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(resim_no, regex=False)]

            if kutup:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(f"-{kutup}", regex=False) | tum_metin.str.contains(f" {kutup} ", regex=False)]

            if govde_boyu:
                gb_bosluksuz = govde_boyu.replace(" ", "").replace("-", "").replace("/", "")
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(gb_bosluksuz, regex=False)]

            if kes_isl:
                kes_isl_temiz = kes_isl.replace("İ", "I").replace("Ş", "S")
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(kes_isl_temiz, regex=False)]

            if mil_boyu:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(mil_boyu, regex=False)]

            if mil_cap:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(mil_cap, regex=False)]

            if revizyon:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"R:{revizyon}", regex=False) |
                    tum_metin.str.contains(f"-{revizyon}-", regex=False) |
                    tum_metin.str.contains(f"R: {revizyon}", regex=False)
                ]

            if ek_aciklama:
                ea_temiz = ek_aciklama.replace(" ", "")
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(ea_temiz, regex=False)]

            if guduk_secili:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains("GUDUK", regex=False)]

            # --- İŞTE DÜZELTİLEN YENİ VERSİYON MANTIĞI ---
            if versiyon == "ESKİ_GÖVDE" or versiyon == "ESKI GOVDE":
                df_filtrelenmis = df_filtrelenmis[
                    df_filtrelenmis.iloc[:, 0].astype(str).str.strip().str.endswith("-00") |
                    df_filtrelenmis.iloc[:, 0].astype(str).str.contains("-00-", regex=False)
                ]
            elif versiyon == "PREMIUM":
                df_filtrelenmis = df_filtrelenmis[
                    ~df_filtrelenmis.iloc[:, 0].astype(str).str.strip().str.endswith("-00") &
                    ~df_filtrelenmis.iloc[:, 0].astype(str).str.contains("-00-", regex=False)
                ]

            if sira_no.isdigit():
                aranan_kod = sira_no.zfill(4)
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(aranan_kod, regex=False)]

            if self.cb_tarih_kullan.isChecked() and len(df_filtrelenmis.columns) > 0:
                b_tar = pd.to_datetime(self.bas_tarih.date().toString("yyyy-MM-dd"))
                b_bit = pd.to_datetime(self.bit_tarih.date().toString("yyyy-MM-dd"))
                parsed_dates = pd.to_datetime(df_filtrelenmis.iloc[:, -1], dayfirst=True, errors='coerce')
                df_filtrelenmis = df_filtrelenmis[(parsed_dates >= b_tar) & (parsed_dates <= b_bit + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

            self.parent.tablo.setRowCount(0)

            for index, row in df_filtrelenmis.iterrows():
                ilk_hucre = str(row.iloc[0]).strip()
                if not ilk_hucre or ilk_hucre == 'nan' or "KOD" in ilk_hucre.upper():
                    continue

                mevcut_satir = self.parent.tablo.rowCount()
                self.parent.tablo.insertRow(mevcut_satir)

                for col_idx in range(min(8, len(row))):
                    veri = row.iloc[col_idx]
                    if pd.notna(veri) and str(veri).strip() != 'nan':
                        self.parent.tablo.setItem(mevcut_satir, col_idx, QTableWidgetItem(str(veri).strip()))

            self.parent.tablo.setSortingEnabled(True)
            self.parent.satir_sayaci.setText(f"Bulunan Kayıt: {self.parent.tablo.rowCount()}")
            self.accept()

        except Exception as e:
            self.parent.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Filtreleme sırasında hata oluştu:\n{e}")