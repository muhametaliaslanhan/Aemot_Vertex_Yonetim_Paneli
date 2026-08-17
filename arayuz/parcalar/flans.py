from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QSpacerItem, QSizePolicy,
    QScrollArea, QMessageBox, QCompleter, QDialog, QDateEdit, QFormLayout,QMenu
)
from PyQt6.QtCore import Qt, QTimer, QDateTime,QDate
from PyQt6.QtGui import QPixmap
import pandas as pd
import re
import os
import stat
import shutil
import tempfile
import xlwings as xw
from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir, kaynak_yolu, katalog_oku
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.motor_verileri import turkce_karakter_temizleme
from arayuz.mantik import oturum

class FlansArayuzu(QWidget):
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
        self.logo_etiketi.setPixmap(QPixmap(kaynak_yolu("arayuz/varliklar/aemot_logo_gorseli.png")))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)

        baslik = QLabel("FLANŞ ARAYÜZÜ")
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

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630","063-090","071-090","100-112","100-132","160-180","180-225","355-400"])

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["ESKİ_GÖVDE", "PREMIUM"])

        self.yapi_sekli_kutusu = QComboBox()
        self.yapi_sekli_kutusu.addItems(["B3", "B5", "B9", "B14", "B34"])

        self.flans_olcusu_kutusu = QComboBox()
        self.flans_olcusu_kutusu.addItems(
            ["YOK", "C90", "C105", "C120", "C140", "C160", "C200", "C250", "A140", "A160", "A200", "A250", "A300",
             "A350", "A400", "A450", "A550", "A660", "A800", "A1000", "A1150", "A1370"])

        duzen_temel.addWidget(QLabel("Versiyon:"), 0, 0)
        duzen_temel.addWidget(self.versiyon_kutusu, 0, 1)

        duzen_temel.addWidget(QLabel("Yapı Şekli):"), 0, 2)
        duzen_temel.addWidget(self.yapi_sekli_kutusu, 0, 3)

        duzen_temel.addWidget(QLabel("Grup (M):"), 1, 0)
        duzen_temel.addWidget(self.grup_kutusu, 1, 1)

        duzen_temel.addWidget(QLabel("Flanş Ölçüsü:"), 1, 2)
        duzen_temel.addWidget(self.flans_olcusu_kutusu, 1, 3)

        grup_temel.setLayout(duzen_temel)
        form_ana_duzen.addWidget(grup_temel)

        ## --- GRUP 2: Klemens Kutu Kapağı Detayları ---
        grup_ozel = QGroupBox("Flanş Detayları")
        duzen_ozel = QGridLayout()

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        # EXCEL FORMÜLÜ İLE UYUMLU OLMASI İÇİN TÜRKÇE KARAKTERLER DÜZELTİLDİ
        self.islenmis_islenmemis_kutusu = QComboBox()
        self.islenmis_islenmemis_kutusu.addItems(["ISLENMIS", "ISLENMEMIS"])

        self.flans_malzemesi_kutusu = QComboBox()
        self.flans_malzemesi_kutusu.addItems(["ALUMINYUM", "PIK", "CELIK"])

        self.resim_no_kutusu = QLineEdit()
        self.ek_aciklama_input = QLineEdit()

        duzen_ozel.addWidget(QLabel("Revizyon No:"), 0, 0)
        duzen_ozel.addWidget(self.revizyon_kutusu, 0, 1)

        duzen_ozel.addWidget(QLabel("Ek Açıklama:"), 0, 2)
        duzen_ozel.addWidget(self.ek_aciklama_input, 0, 3)

        duzen_ozel.addWidget(QLabel("İşlenmiş/İşlenmemiş:"), 1, 0)
        duzen_ozel.addWidget(self.islenmis_islenmemis_kutusu, 1, 1)

        duzen_ozel.addWidget(QLabel("Gövde Malzemesi:"), 1, 2)
        duzen_ozel.addWidget(self.flans_malzemesi_kutusu, 1, 3)

        duzen_ozel.addWidget(QLabel("Resim No:"), 2, 0)
        duzen_ozel.addWidget(self.resim_no_kutusu, 2, 1)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        # --- GRUP 3: Ek Özellikler (Onay Kutuları) ---
        grup_onaylar = QGroupBox("Ek Donanım ve Özellikler")
        duzen_onaylar = QGridLayout()

        self.cb_frenli = QCheckBox("FRENLİ")
        self.cb_derin = QCheckBox("DERİN")
        self.cb_burclu = QCheckBox("BURCLU")
        self.cb_guduk = QCheckBox("GÜDÜK")
        self.cb_segmanli = QCheckBox("SEGMANLI")
        self.cb_yag_keceli = QCheckBox("YAG KECELI")
        self.cb_drenaj_tapali = QCheckBox("DRENAJ TAPALI")
        self.cb_yaglamali = QCheckBox("YAGLAMALI")

        # 1. Satır
        duzen_onaylar.addWidget(self.cb_frenli, 0, 0)
        duzen_onaylar.addWidget(self.cb_derin, 0, 1)
        duzen_onaylar.addWidget(self.cb_burclu, 0, 2)
        duzen_onaylar.addWidget(self.cb_guduk, 0, 3)

        # 2. Satır
        duzen_onaylar.addWidget(self.cb_segmanli, 1, 0)
        duzen_onaylar.addWidget(self.cb_yag_keceli, 1, 1)
        duzen_onaylar.addWidget(self.cb_drenaj_tapali, 1, 2)
        duzen_onaylar.addWidget(self.cb_yaglamali, 1, 3)

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
        self.tablo = QTableWidget(0, 9)  # DİKKAT: 8 YERİNE 9 OLDU!
        self.tablo_basliklari = ["#KOD", "AÇIKLAMA", "GRUP", 'AÇIKLAMA2', "RESİM NO", 'AÇIKLAMA SIRA NO',
                                 'AÇIKLAMA SIRA KOD', "Sütun1", "KAYIT TARİHİ"]
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
        self.tablo.setAlternatingRowColors(True)
        ana_duzen.addWidget(self.tablo, stretch=1)

        # ---------------- EXCEL'E AKTAR BUTONU ----------------
        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR (FLANŞ)")
        self.aktar_butonu.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.aktar_butonu.clicked.connect(self.excele_aktar)
        ana_duzen.addWidget(self.aktar_butonu)

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
            nesne.clear()
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

        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = "FLANŞ"
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.strip().upper() == "FLANŞ":
                    hedef_sayfa = sayfa
                    break

            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]

            for index, row in df.iterrows():
                val_str = str(row.iloc[0]).strip().upper()
                aciklama_str = str(row.iloc[1]).strip()

                if "-FL-" in val_str and aciklama_str:
                    parcalar = val_str.split("-")
                    if len(parcalar) >= 3:
                        sira_kodu_kismi = parcalar[2]
                        if len(sira_kodu_kismi) >= 4 and sira_kodu_kismi.isdigit():
                            kullanilan_nolar.add(int(sira_kodu_kismi))
        except Exception as e:
            pass

        for row_idx in range(self.tablo.rowCount()):
            try:
                tablo_id = self.tablo.item(row_idx, 0).text().strip().upper()
                if "-FL-" in tablo_id:
                    parcalar = tablo_id.split("-")
                    if len(parcalar) >= 3:
                        sira_kodu_kismi = parcalar[2]
                        if len(sira_kodu_kismi) >= 4 and sira_kodu_kismi.isdigit():
                            kullanilan_nolar.add(int(sira_kodu_kismi))
            except:
                pass

        for i in range(1, 10000):
            if i not in kullanilan_nolar:
                return i

        return len(kullanilan_nolar) + 1

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip()).upper()

        guncel_yol = ayar_excel_yolunu_getir()
        self.tablo.setRowCount(0)

        gercek_sayfa_adi = "FLANŞ"
        try:
            excel_dosyasi = pd.ExcelFile(guncel_yol)
            for sayfa in excel_dosyasi.sheet_names:
                temiz_sayfa_adi = sayfa.strip().upper()
                if temiz_sayfa_adi == "FLANŞ":
                    gercek_sayfa_adi = sayfa
                    break
        except:
            pass

        try:
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=gercek_sayfa_adi, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]
            df_veri = df.copy()
            bulunan_kayit_sayisi = 0
            self.tablo.setSortingEnabled(False)

            for index, row in df_veri.iterrows():
                ilk_hucre = str(row.iloc[0]).strip()
                if not ilk_hucre or ilk_hucre == 'nan' or "KOD" in ilk_hucre.upper():
                    continue

                tum_satir = " ".join([str(val).strip() for val in row.values if pd.notna(val)])
                tam_havuz = turkce_karakter_temizleme(tum_satir).upper()
                metin_sifir_bosluk = tam_havuz.replace(" ", "").replace("-", "").replace("/", "")

                id_eslesti = True
                aciklama_eslesti = True

                if aranan_id:
                    aranan_id_temiz = aranan_id.replace(" ", "").replace("-", "")
                    if aranan_id_temiz not in metin_sifir_bosluk:
                        id_eslesti = False

                if aranan_aciklama:
                    aranan_kelimeler = aranan_aciklama.split()
                    if not all(kelime in tam_havuz for kelime in aranan_kelimeler):
                        aciklama_eslesti = False

                if id_eslesti and aciklama_eslesti:
                    mevcut_satir = self.tablo.rowCount()
                    self.tablo.insertRow(mevcut_satir)

                    tablo_sutun_sayisi = self.tablo.columnCount()
                    for col_idx in range(min(tablo_sutun_sayisi, len(row))):
                        veri = row.iloc[col_idx]
                        if pd.notna(veri) and str(veri).strip() != 'nan':
                            self.tablo.setItem(mevcut_satir, col_idx, QTableWidgetItem(str(veri).strip()))

                    bulunan_kayit_sayisi += 1

            self.tablo.setSortingEnabled(True)
            self.satir_sayaci.setText(f"Bulunan Kayıt: {bulunan_kayit_sayisi}")

            if bulunan_kayit_sayisi == 0:
                QMessageBox.information(self, "Bulunamadı", "Aradığınız kriterlerde kayıt bulunamadı.")

        except ValueError:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Excel dosyasında '{gercek_sayfa_adi}' sayfası bulunamadı!")
        except Exception as e:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Arama sırasında hata:\n{e}")

    def gelismis_arama_ac(self):
        try:
            excel_yolu = ayar_excel_yolunu_getir()
            dialog = FlansGelismisArama(self, excel_yolu)
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Gelişmiş arama açılamadı:\n{e}")

    def excele_aktar(self):
        gercek_sayfa_adi = "FLANŞ"
        guncel_yol = ayar_excel_yolunu_getir()
        excel_yolu = os.path.abspath(guncel_yol)

        try:
            excel_dosyasi = pd.ExcelFile(excel_yolu)
            for sayfa in excel_dosyasi.sheet_names:
                temiz_sayfa = sayfa.upper().strip()
                if temiz_sayfa == "FLANŞ":
                    gercek_sayfa_adi = sayfa
                    break
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Excel okunamadı:\n{e}")
            return

        if self.tablo.rowCount() == 0:
            QMessageBox.warning(self, "Uyarı", "Tabloda aktarılacak veri yok!")
            return

        # 1. KOPYA KONTROLÜ
        try:
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=gercek_sayfa_adi, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]

            # --- DEĞİŞEN KISIM BURASI: Artık KOD değil AÇIKLAMA arıyoruz ---
            aciklama_kolonu = None
            for col in df.columns:
                if "AÇIKLAMA" == col or "ACIKLAMA" == col:
                    aciklama_kolonu = col
                    break

            exceldeki_aciklamalar = df[aciklama_kolonu].astype(str).str.strip().tolist() if aciklama_kolonu else []
            cakisan_satirlar = []

            for satir in range(self.tablo.rowCount()):
                # DİKKAT: 0 (KOD) yerine 1 (AÇIKLAMA) sütunundaki veriyi çekiyoruz!
                item = self.tablo.item(satir, 1)
                if item and item.text().strip() in exceldeki_aciklamalar:
                    cakisan_satirlar.append(satir)
            # ----------------------------------------------------------------

            if cakisan_satirlar:
                satir_numaralari = ", ".join([str(r + 1) for r in cakisan_satirlar])
                cevap_kopya = QMessageBox.question(
                    self, "KOPYA FLANŞ UYARISI!",
                    f"Tablodaki {satir_numaralari}. satırdaki flanş(lar)ın AÇIKLAMASI Excel'de zaten mevcut!\n\n"
                    f"• EVET: Kopyaları tablodan kaldır ve sadece YENİLERİ kaydet.\n"
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

        cevap = QMessageBox.question(self, "ONAY", f"{gercek_sayfa_adi} sayfasına aktarılsın mı?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if cevap == QMessageBox.StandardButton.No: return

        # 3. GÖLGE KOPYA (SHADOW COPY) VE YAZMA SİSTEMİ
        try:
            temp_klasor = tempfile.gettempdir()
            gecici_excel = os.path.join(temp_klasor, "gecici_aemot_flans.xlsx")
            shutil.copy2(excel_yolu, gecici_excel)
            os.chmod(gecici_excel, stat.S_IWRITE)

            app = xw.App(visible=False)
            app.display_alerts = False

            wb = app.books.open(gecici_excel)
            sheet = wb.sheets[gercek_sayfa_adi]

            # DİKKAT: Tablomuz 9 Sütun! Kayıt tarihi I8'e yazılır.
            # --- 1. Satıra başlıkları yazdır ve A1 formatına uyarla ---
            sheet.range('I1').value = "KAYIT TARİHİ"
            sheet.range('J1').value = "KAYDI EKLEYEN KİŞİ"
            try:
                sheet.range('A1:J1').api.Font.Bold = True
            except:
                pass

            son_satir = sheet.range('A1048576').end('up').row
            if son_satir < 1: son_satir = 1

            aktarilacak_veriler = []
            for r in range(self.tablo.rowCount()):
                satir_verisi = []
                for c in range(9):  # 9 SÜTUN
                    item = self.tablo.item(r, c)
                    satir_verisi.append(item.text() if item else "")

                # SATIRIN SONUNA KULLANICIYI ZIMBALA
                satir_verisi.append(oturum.tam_isim())
                aktarilacak_veriler.append(satir_verisi)

            if aktarilacak_veriler:
                baslama_satiri = son_satir + 1
                bitis_satiri = baslama_satiri + len(aktarilacak_veriler) - 1

                # J sütununa kadar yazdır
                hedef_alan = sheet.range(f'A{baslama_satiri}:J{bitis_satiri}')
                hedef_alan.number_format = '@'
                hedef_alan.value = aktarilacak_veriler

            wb.save(gecici_excel)
            wb.close()
            app.quit()

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

    def dna_bazli_sira_no_bul(self, sayfa_adi, aranan_aciklama2):
        hedef_dna = turkce_karakter_temizleme(aranan_aciklama2).replace(" ", "").upper()

        # --- 1. ARAYÜZ TABLOSU ---
        for row_idx in range(self.tablo.rowCount()):
            tablo_aciklama2 = turkce_karakter_temizleme(str(self.tablo.item(row_idx, 3).text())).replace(" ",
                                                                                                         "").upper()

            if tablo_aciklama2 == hedef_dna:
                return int(self.tablo.item(row_idx, 5).text().strip())

        # --- 2. EXCEL GEÇMİŞİ ---
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = "FLANŞ"
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.strip().upper() == "FLANŞ":
                    hedef_sayfa = sayfa
                    break

            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]

            for index, row in df.iterrows():
                if len(row) > 5:
                    kod_hucre = str(row.iloc[0]).strip().upper()
                    if "-FL-" not in kod_hucre:
                        continue

                    # 3. İndeks = AÇIKLAMA 2 | 4. İndeks = RESİM NO | 5. İndeks = AÇIKLAMA S.NO
                    excel_aciklama2 = turkce_karakter_temizleme(str(row.iloc[3])).replace(" ", "").upper()

                    if excel_aciklama2 == hedef_dna:
                        sira_no = str(row.iloc[5]).replace("'", "").strip()
                        try:
                            return int(float(sira_no))
                        except:
                            pass
        except Exception as e:
            pass

        return None

    def uret_calistir(self):
        for kutu in self.findChildren(QComboBox):
            yeni_metin = kutu.currentText().strip().upper()
            if yeni_metin and kutu.findText(yeni_metin) == -1:
                kutu.addItem(yeni_metin)

        # 1. ARAYÜZDEN VERİLERİ ÇEK
        versiyon = self.versiyon_kutusu.currentText().strip()
        grup = self.grup_kutusu.currentText().strip()
        yapi_sekli = self.yapi_sekli_kutusu.currentText().strip()
        flans_olcusu = self.flans_olcusu_kutusu.currentText().strip()
        revizyon = self.revizyon_kutusu.currentText().strip()
        islenmis = self.islenmis_islenmemis_kutusu.currentText().strip()
        malzeme = self.flans_malzemesi_kutusu.currentText().strip()
        resim_no = self.resim_no_kutusu.text().strip()

        ham_ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
        frenli_mi = self.cb_frenli.isChecked()

        # --- YENİ EKLENEN: TÜM GİRİLEBİLİR ALANLAR İÇİN BOŞLUK UYARISI ---
        bos_alanlar = []
        if not versiyon: bos_alanlar.append("Versiyon")
        if not grup: bos_alanlar.append("Grup")
        if not islenmis: bos_alanlar.append("İşlenmiş Durumu")
        if not malzeme: bos_alanlar.append("Malzeme")
        if not revizyon: bos_alanlar.append("Revizyon No")
        if not resim_no: bos_alanlar.append("Resim No")

        # Frenli değilse Yapı Şekli ve Flanş Ölçüsü uyarısı
        if not frenli_mi:
            if not yapi_sekli: bos_alanlar.append("Yapı Şekli")
            if not flans_olcusu: bos_alanlar.append("Flanş Ölçüsü")

        if bos_alanlar:
            cevap = QMessageBox.question(self, "Eksik Bilgi",
                                         f"Şu alanlar boş bırakıldı:\n{', '.join(bos_alanlar)}\n\nYine de üretip tabloya eklemek istiyor musunuz?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        if not resim_no:
            resim_no = "0000"

        # ---------------------------------------------------------
        # NOKTA ATIŞI STANDART FLANŞ KONTROLÜ VE ONAY
        # ---------------------------------------------------------
        yapi_kontrol = "B14" if frenli_mi else yapi_sekli.upper()
        flans_kontrol = flans_olcusu.upper()

        beklenen_flans = ""
        alternatifler = []

        try:
            import os, json
            katalog_yolu = os.path.join(os.getcwd(), "arayuz", "mantik", "katalog_veri_json", "standart_flans.json")
            if not os.path.exists(katalog_yolu):
                katalog_yolu = os.path.join(os.getcwd(), "mantik", "katalog_veri_json", "standart_flans.json")

            with open(katalog_yolu, "r", encoding="utf-8") as f:
                std_flans_katalogu = json.load(f)

            for key, val in std_flans_katalogu.items():
                if isinstance(val, dict):
                    m_json = str(val.get("M", "")).strip().upper()
                    yapi_json = str(val.get("Yapi Sekli", "")).strip().upper()
                    flans_json = str(val.get("flans olcusu", "")).strip().upper()

                    if m_json.lstrip("0") == grup.lstrip("0"):
                        if yapi_json in ["B5", "B14", "B34"]:
                            alternatifler.append(f"• Yapı Şekli: {yapi_json} -> Flanş: {flans_json}")
                        if yapi_json == yapi_kontrol:
                            beklenen_flans = flans_json
        except Exception as e:
            pass

        # UYARI VE ONAY MEKANİZMASI
        if yapi_kontrol in ["B5", "B14", "B34"]:
            if not beklenen_flans:
                alt_metin = "\n".join(sorted(list(
                    set(alternatifler)))) if alternatifler else f"Bu gruba ({grup}) ait tanımlı standart flanş kaydı yok."
                cevap = QMessageBox.question(self, "STANDART FLANŞ BULUNAMADI!",
                                             f"Seçtiğiniz Gövde ({grup}) ve Yapı Şekli ({yapi_kontrol}) için sistemde tanımlı bir flanş değeri YOK!\n\n"
                                             f"Aynı gruba ({grup}) ait mevcut standartlar şunlardır:\n\n{alt_metin}\n\n"
                                             f"Yine de flanş ölçüsünü '{flans_kontrol}' olarak ayarlayıp kod üretmek istiyor musunuz?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if cevap == QMessageBox.StandardButton.No: return

            elif beklenen_flans and flans_kontrol != beklenen_flans:
                cevap = QMessageBox.question(self, "STANDART FLANŞ UYUMSUZLUĞU!",
                                             f"Girdiğiniz flanş ölçüsü hatalı veya standart dışı!\n\n"
                                             f"Gövde (M): {grup}\nYapı Şekli: {yapi_kontrol}\nSizin Seçiminiz: {flans_kontrol if flans_kontrol and flans_kontrol != 'YOK' else 'GİRİLMEDİ (YOK)'}\n\n"
                                             f"Olması Gereken Standart Flanş: {beklenen_flans}\n\n"
                                             f"Yine de flanşı '{flans_kontrol}' olarak kabul edip üretime devam etmek istiyor musunuz?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if cevap == QMessageBox.StandardButton.No: return

        # Checkboxların Eklenmesi (Eski koddaki yerinden aynen devam eder)
        ek_ozellikler = []
        if getattr(self, "cb_derin", None) and self.cb_derin.isChecked(): ek_ozellikler.append("DERIN")
        if self.cb_burclu.isChecked(): ek_ozellikler.append("BURCLU")
        if self.cb_guduk.isChecked(): ek_ozellikler.append("GUDUK")
        if self.cb_segmanli.isChecked(): ek_ozellikler.append("SEGMANLI")
        if self.cb_yag_keceli.isChecked(): ek_ozellikler.append("YAG KECELI")
        if self.cb_drenaj_tapali.isChecked(): ek_ozellikler.append("DRENAJ TAPALI")
        if self.cb_yaglamali.isChecked(): ek_ozellikler.append("YAGLAMALI")

        for ozellik in ek_ozellikler:
            if ozellik not in ham_ek_aciklama.replace("Ü", "U"):
                ham_ek_aciklama += f" {ozellik}"

        ek_aciklama = ham_ek_aciklama.strip()

        if "FREN" in ek_aciklama:
            QMessageBox.critical(self, "KURAL İHLALİ", "EK AÇIKLAMADA FREN TANIMI OLAMAZ.")
            return

        # ---------------- 2. EXCEL FORMÜLÜNE GÖRE TASLAK AÇIKLAMA OLUŞTURMA ----------------
        taslak_parcalari = [grup]

        if frenli_mi:
            taslak_parcalari.append("B14")
        else:
            if yapi_sekli:
                taslak_parcalari.append(yapi_sekli)
            if flans_olcusu and flans_olcusu != "YOK":
                taslak_parcalari.append(flans_olcusu)

        if ek_aciklama:
            taslak_parcalari.append(ek_aciklama)

        if malzeme:
            taslak_parcalari.append(malzeme)

        if frenli_mi:
            taslak_parcalari.append("FREN")

        if islenmis:
            taslak_parcalari.append(islenmis)

        taslak_parcalari.append("FLANS")

        if versiyon != "ESKİ_GÖVDE":
            taslak_parcalari.append(f"- R:{revizyon}")

        if resim_no and resim_no != "0000":
            taslak_parcalari.append(f"- {resim_no}")

        taslak_aciklama = " ".join(taslak_parcalari)
        taslak_aciklama = " ".join(taslak_aciklama.split())

        # =========================================================================
        # --- AÇIKLAMA TABANLI KOPYA KALKANI ---
        # =========================================================================
        for r in range(self.tablo.rowCount()):
            tablo_aciklama = self.tablo.item(r, 1).text().strip()
            if tablo_aciklama == taslak_aciklama:
                QMessageBox.warning(self, "KOPYA ENGELİ",
                                    f"Bu AÇIKLAMAYA sahip Flanş ZATEN LİSTEDE var!\n\nAynı özellikleri ard arda ekleyemezsiniz.")
                return

        # --- AÇIKLAMA 2'Yİ OTOMATİK ÜRETMESİ ---
        # --- AÇIKLAMA 2'Yİ OTOMATİK ÜRETMESİ ---
        aciklama2_oto = taslak_aciklama

        # 1. Grubu baştan sil
        if aciklama2_oto.startswith(grup):
            aciklama2_oto = aciklama2_oto[len(grup):].strip()

        # 2. Resim numarasını sondan sil (Sadece 0000'dan farklıysa silinir)
        if resim_no and resim_no != "0000":
            aciklama2_oto = aciklama2_oto.replace(f"- {resim_no}", "").strip()

        # NOT: Revizyona (R:01 vs.) hiç dokunmuyoruz!
        # Eski gövde ise zaten metinde yok, Premium ise metinde olduğu gibi kalacak.

        # ---------------- 3. DNA TESTİ İLE SIRA NO BULMA ----------------
        dna_sira = self.dna_bazli_sira_no_bul(sayfa_adi="FLANŞ", aranan_aciklama2=aciklama2_oto)

        from arayuz.mantik.sira_no_dialog import sira_no_secim_yap
        secilen_sira = sira_no_secim_yap(self, dna_sira, aciklama2_oto, sayfa_adi="FLANŞ", id_kodu="FL")
        if secilen_sira is None:
            return
        elif secilen_sira == -1:
            sira_no_int = self.genel_max_sira_no_bul()  # DÜZELTME: fonksiyon zaten ilk boş sıra noyu döndürüyor, +1 eklemek onu dolu bir numaraya kaydırıyordu
        else:
            sira_no_int = secilen_sira

        sira_no_str = str(sira_no_int)
        aciklama_kod_str = sira_no_str.zfill(4)

        # ---------------- 4. NİHAİ ID OLUŞTURMA ----------------
        id_grup = grup[:3]
        id_revizyon = "00" if versiyon == "ESKİ_GÖVDE" else revizyon
        id_resim = "0000" if not resim_no else resim_no.zfill(4)

        uretilen_id = f"YMI{id_grup}-FL-{aciklama_kod_str}-{id_revizyon}-{id_resim}"

        # ---------------- 5. TABLOYA YAZDIRMA ----------------
        bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

        # DİKKAT: 9 Sütunlu yapı (Sütun1 için "" boş string eklendi, Kayıt Tarihi sona atıldı)
        yeni_kayit = [
            uretilen_id, taslak_aciklama, grup, aciklama2_oto, resim_no, sira_no_str, aciklama_kod_str, "",
            bugunun_tarihi
        ]

        self.tablo.setSortingEnabled(False)
        mevcut_satir = self.tablo.rowCount()
        self.tablo.insertRow(mevcut_satir)
        for sutun, veri in enumerate(yeni_kayit):
            hucre = QTableWidgetItem(str(veri))
            self.tablo.setItem(mevcut_satir, sutun, hucre)

        self.tablo.setSortingEnabled(True)

    def tablo_sag_tik(self, pos):

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
        from PyQt6.QtWidgets import QMenu
        menu = QMenu()

        aksiyon_sutun_sigdir = menu.addAction("Bu Sütunu İçeriğe Sığdır")
        aksiyon_tum_sutunlar = menu.addAction("Tüm Sütunları İçeriğe Sığdır")

        sutun_index = self.tablo.horizontalHeader().logicalIndexAt(pos)
        secim = menu.exec(self.tablo.horizontalHeader().mapToGlobal(pos))

        if secim == aksiyon_sutun_sigdir:
            self.tablo.resizeColumnToContents(sutun_index)
        elif secim == aksiyon_tum_sutunlar:
            self.tablo.resizeColumnsToContents()


class FlansGelismisArama(QDialog):
    def __init__(self, parent, excel_yolu):
        super().__init__(parent)
        self.setWindowTitle("FLANŞ Gelişmiş Arama")
        self.setMinimumWidth(850)
        self.parent = parent
        self.excel_yolu = excel_yolu

        self.layout = QVBoxLayout(self)

        # --- 1. EXCEL VERİSİNİ OKU ---
        gercek_sayfa = "FLANŞ"
        try:
            dosya = pd.ExcelFile(self.excel_yolu)
            for s in dosya.sheet_names:
                if s.strip().upper() == "FLANŞ":
                    gercek_sayfa = s
                    break
            self.df = pd.read_excel(self.excel_yolu, sheet_name=gercek_sayfa, header=0, dtype=str).fillna("")
            self.df.columns = [str(c).strip().upper() for c in self.df.columns]
        except:
            self.df = pd.DataFrame()

        # --- 2. FORM ALANI ---
        grup_form = QGroupBox("FLANŞ Parametreleriyle Filtrele")
        form_layout = QGridLayout()

        # Kutular
        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["", "ESKİ_GÖVDE", "PREMIUM"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["", "063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355", "400", "450", "500", "630"])

        self.yapi_sekli_kutusu = QComboBox()
        self.yapi_sekli_kutusu.addItems(["", "B3", "B5", "B9", "B14", "B34"])

        self.flans_olcusu_kutusu = QComboBox()
        self.flans_olcusu_kutusu.addItems(
            ["", "YOK", "C90", "C105", "C120", "C140", "C160", "C200", "C250", "A140", "A160", "A200", "A250", "A300",
             "A350", "A400", "A450", "A550", "A660", "A800", "A1000", "A1150", "A1370"])

        self.islenmis_islenmemis_kutusu = QComboBox()
        self.islenmis_islenmemis_kutusu.addItems(["", "ISLENMIS", "ISLENMEMIS"])

        self.flans_malzemesi_kutusu = QComboBox()
        self.flans_malzemesi_kutusu.addItems(["", "ALUMINYUM", "PIK", "CELIK"])

        self.resim_no_input = QLineEdit()
        self.resim_no_input.setPlaceholderText("Örn: 3130")

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.setEditable(True)
        self.revizyon_kutusu.addItems(["", "01", "02", "03", "04", "05"])

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Ek Açıklama")

        self.sira_no_kutusu = QLineEdit()
        self.sira_no_kutusu.setPlaceholderText("Örn: 1")

        # Yerleştirme
        form_layout.addWidget(QLabel("Versiyon:"), 0, 0)
        form_layout.addWidget(self.versiyon_kutusu, 0, 1)
        form_layout.addWidget(QLabel("Grup (M):"), 0, 2)
        form_layout.addWidget(self.grup_kutusu, 0, 3)

        form_layout.addWidget(QLabel("Yapı Şekli:"), 1, 0)
        form_layout.addWidget(self.yapi_sekli_kutusu, 1, 1)
        form_layout.addWidget(QLabel("Flanş Ölçüsü:"), 1, 2)
        form_layout.addWidget(self.flans_olcusu_kutusu, 1, 3)

        form_layout.addWidget(QLabel("İşlenmiş Durumu:"), 2, 0)
        form_layout.addWidget(self.islenmis_islenmemis_kutusu, 2, 1)
        form_layout.addWidget(QLabel("Malzeme:"), 2, 2)
        form_layout.addWidget(self.flans_malzemesi_kutusu, 2, 3)

        form_layout.addWidget(QLabel("Revizyon No:"), 3, 0)
        form_layout.addWidget(self.revizyon_kutusu, 3, 1)
        form_layout.addWidget(QLabel("Resim No:"), 3, 2)
        form_layout.addWidget(self.resim_no_input, 3, 3)

        form_layout.addWidget(QLabel("Ek Açıklama:"), 4, 0)
        form_layout.addWidget(self.ek_aciklama_input, 4, 1)
        form_layout.addWidget(QLabel("Sıra No (Kod):"), 4, 2)
        form_layout.addWidget(self.sira_no_kutusu, 4, 3)

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

        # --- BUTON ---
        self.btn_ara = QPushButton("Filtrele ve Tabloya Getir")
        self.btn_ara.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
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

            # --- 1. DNA / Metin Havuzu Oluşturma ---
            tum_metin = pd.Series([""] * len(df_filtrelenmis), index=df_filtrelenmis.index)
            for col_idx in range(min(8, len(df_filtrelenmis.columns))):
                tum_metin += df_filtrelenmis.iloc[:, col_idx].astype(str).str.upper() + " "

            tum_metin = tum_metin.str.replace("İ", "I").str.replace("Ş", "S").str.replace("Ğ", "G").str.replace("Ü", "U").str.replace("Ö", "O").str.replace("Ç", "C")
            metin_sifir_bosluk = tum_metin.str.replace(" ", "").str.replace("-", "").str.replace("/", "")

            # Form Girdilerini Çek (FLANŞ)
            versiyon = self.versiyon_kutusu.currentText().strip().upper()
            grup = self.grup_kutusu.currentText().strip().upper()
            yapi_sekli = self.yapi_sekli_kutusu.currentText().strip().upper()
            flans_olcusu = self.flans_olcusu_kutusu.currentText().strip().upper()
            islenmis = self.islenmis_islenmemis_kutusu.currentText().strip().upper()
            malzeme = self.flans_malzemesi_kutusu.currentText().strip().upper()
            resim_no = self.resim_no_input.text().strip().upper()
            revizyon = self.revizyon_kutusu.currentText().strip().upper()
            ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
            sira_no = self.sira_no_kutusu.text().strip()

            # --- AKILLI FİLTRELEME KURALLARI ---
            if grup and len(df_filtrelenmis.columns) > 2:
                df_filtrelenmis = df_filtrelenmis[df_filtrelenmis.iloc[:, 2].astype(str).str.strip().str.upper() == grup]

            if yapi_sekli:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(yapi_sekli, regex=False)]

            if flans_olcusu:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(flans_olcusu, regex=False)]

            if islenmis:
                islenmis_temiz = islenmis.replace("İ", "I").replace("Ş", "S")
                if islenmis_temiz == "ISLENMIS":
                    df_filtrelenmis = df_filtrelenmis[
                        metin_sifir_bosluk.str.contains("ISLENMIS", regex=False) &
                        ~metin_sifir_bosluk.str.contains("ISLENMEMIS", regex=False)
                    ]
                else:
                    df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(islenmis_temiz, regex=False)]

            if malzeme:
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(malzeme, regex=False)]

            if resim_no:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(resim_no, regex=False)]

            if revizyon:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"R:{revizyon}", regex=False) | tum_metin.str.contains(f"-{revizyon}", regex=False)]

            if ek_aciklama:
                ea_temiz = ek_aciklama.replace(" ", "")
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(ea_temiz, regex=False)]

            # EVRENSEL VERSİYON (PREMIUM / ESKİ GÖVDE)
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
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(aranan_kod, regex=False)]

            if self.cb_tarih_kullan.isChecked() and len(df_filtrelenmis.columns) > 0:
                b_tar = pd.to_datetime(self.bas_tarih.date().toString("yyyy-MM-dd"))
                b_bit = pd.to_datetime(self.bit_tarih.date().toString("yyyy-MM-dd"))
                parsed_dates = pd.to_datetime(df_filtrelenmis.iloc[:, -1], dayfirst=True, errors='coerce')
                df_filtrelenmis = df_filtrelenmis[
                    (parsed_dates >= b_tar) & (parsed_dates <= b_bit + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

            # --- TABLOYA AKTARIM ---
            self.parent.tablo.setRowCount(0)
            tablo_sutun_sayisi = self.parent.tablo.columnCount()

            for index, row in df_filtrelenmis.iterrows():
                ilk_hucre = str(row.iloc[0]).strip()
                if not ilk_hucre or ilk_hucre == 'nan' or "KOD" in ilk_hucre.upper():
                    continue

                mevcut_satir = self.parent.tablo.rowCount()
                self.parent.tablo.insertRow(mevcut_satir)

                for col_idx in range(min(tablo_sutun_sayisi, len(row))):
                    veri = row.iloc[col_idx]
                    if pd.notna(veri) and str(veri).strip() != 'nan':
                        self.parent.tablo.setItem(mevcut_satir, col_idx, QTableWidgetItem(str(veri).strip()))

            self.parent.tablo.setSortingEnabled(True)
            self.parent.satir_sayaci.setText(f"Bulunan Kayıt: {self.parent.tablo.rowCount()}")
            self.accept()

        except Exception as e:
            self.parent.tablo.setSortingEnabled(True)
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Hata", f"Filtreleme sırasında hata oluştu:\n{e}")