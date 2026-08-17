from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QSpacerItem, QSizePolicy,
    QScrollArea, QMessageBox, QCompleter, QDialog, QDateEdit, QFormLayout,QMenu
)
from PyQt6.QtCore import Qt, QTimer, QDateTime,QDate
from PyQt6.QtGui import QPixmap
import pandas as pd
import traceback

import re
import os
import stat
import shutil
import tempfile
import xlwings as xw
from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir, kaynak_yolu
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.motor_verileri import turkce_karakter_temizleme
from arayuz.mantik.json_islemleri import katalog_oku
from arayuz.mantik import oturum




class StatorluGovdeArayuzu(QWidget):
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

        baslik = QLabel("STATORLU GÖVDE")
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

        """VERSİYON	VERİMLİLİK	GRUP	GÖVDE BOYU	KUTUP	PAKET BOYU """

        grup_temel = QGroupBox("Temel STATORLU GOVDE Özellikleri")
        duzen_temel = QGridLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["Eski Govde", "Premium Govde"])

        self.verimlilik_kutusu = QComboBox()
        self.verimlilik_kutusu.addItems(["","IE1", "IE2", "IE3", "IE4", "IE5"])

        self.M_kutusu = QComboBox()
        self.M_kutusu.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(["S", "M", "L"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(
            ["2", "4", "6", "8", "10", "12", "2/4", "4/2", "4/8", "6/4", "6/8", "8/2", "8/4", "8/6", "2/12", "4/12",
             "12/4", "4/16"])

        self.paket_boyu_kutusu = QComboBox()
        self.paket_boyu_kutusu.addItems(["A", "B", "C", "D","K", "P","X", "Y", "Z", "Q", "A-V"])

        duzen_temel.addWidget(QLabel("Versiyon:"), 0, 0)
        duzen_temel.addWidget(self.versiyon_kutusu, 0, 1)

        duzen_temel.addWidget(QLabel("Verimlilik:"), 0, 2)
        duzen_temel.addWidget(self.verimlilik_kutusu, 0, 3)

        duzen_temel.addWidget(QLabel("Gövde:"), 1, 0)
        duzen_temel.addWidget(self.M_kutusu, 1, 1)

        duzen_temel.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        duzen_temel.addWidget(self.govde_boyu_kutusu, 1, 3)

        duzen_temel.addWidget(QLabel("Kutup:"), 2, 0)
        duzen_temel.addWidget(self.kutup_kutusu, 2, 1)

        duzen_temel.addWidget(QLabel("Paket Boyu:"), 2, 2)
        duzen_temel.addWidget(self.paket_boyu_kutusu, 2, 3)

        grup_temel.setLayout(duzen_temel)
        form_ana_duzen.addWidget(grup_temel)

        # --- GRUP 2: Opsiyonel ve Serbest Girişler ---
        """AYAK	GÖVDE MALZEME CİNSİ	REV.NO	RESİM NO EK AÇIKLAMA """

        grup_opsiyon = QGroupBox("Opsiyonel Özellikler")
        duzen_opsiyon = QGridLayout()

        self.ayak_kutusu = QComboBox()
        self.ayak_kutusu.addItems(["AYAKLI", "AYAKSIZ"])

        self.govde_malzemesi_kutusu = QComboBox()
        self.govde_malzemesi_kutusu.addItems(["ALUMINYUM", "PIK", "CELIK"])

        self.resim_no_kutusu = QLineEdit()

        self.ek_aciklama_kutusu = QLineEdit()
        self.ek_aciklama_kutusu.setPlaceholderText("Ek açıklama giriniz...")


        self.ozel_frekans_kutusu = QComboBox()
        self.ozel_frekans_kutusu.addItems(["","33", "50", "60", "67", "75", "87", "100", "106", "120", "133", "135.33"])

        self.ozel_gerilim_kutusu = QLineEdit()
        self.ozel_gerilim_kutusu.setPlaceholderText("Varsa özel gerilim")

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        duzen_opsiyon.addWidget(QLabel("Gövde Malzemesi:"), 0, 0)
        duzen_opsiyon.addWidget(self.govde_malzemesi_kutusu, 0, 1)

        duzen_opsiyon.addWidget(QLabel("Resim No:"), 3, 2 )
        duzen_opsiyon.addWidget(self.resim_no_kutusu, 3, 3)

        duzen_opsiyon.addWidget(QLabel("Ek Açıklama:"), 1, 0)
        duzen_opsiyon.addWidget(self.ek_aciklama_kutusu, 1, 1, 1, 3)  # 3 sütun kaplasın geniş olsun

        duzen_opsiyon.addWidget(QLabel("Revizyon No:"), 2, 0)
        duzen_opsiyon.addWidget(self.revizyon_kutusu, 2, 1)

        duzen_opsiyon.addWidget(QLabel("Özel Frekans:"), 2, 2)
        duzen_opsiyon.addWidget(self.ozel_frekans_kutusu, 2, 3)

        duzen_opsiyon.addWidget(QLabel("Özel Gerilim:"), 3, 0)
        duzen_opsiyon.addWidget(self.ozel_gerilim_kutusu, 3, 1)

        duzen_opsiyon.addWidget(QLabel("Ayak:"),0, 2)
        duzen_opsiyon.addWidget(self.ayak_kutusu, 0, 3)


        grup_opsiyon.setLayout(duzen_opsiyon)
        form_ana_duzen.addWidget(grup_opsiyon)

        # --- GRUP 3: Boya, Özel KW ve Özel Devir ---

        grup_ozel = QGroupBox("Özel KW/DEVİR")
        duzen_ozel = QGridLayout()

        self.cb_ozel_kw = QCheckBox("Özel KW")
        self.le_ozel_kw = QLineEdit()
        self.le_ozel_kw.setEnabled(False)

        self.cb_ozel_devir = QCheckBox("Özel Devir")
        self.le_ozel_devir = QLineEdit()
        self.le_ozel_devir.setEnabled(False)

        # İki checkbox'ın birbirini tetikleme mantığı
        self.cb_ozel_kw.toggled.connect(self.cb_ozel_devir.setChecked)
        self.cb_ozel_devir.toggled.connect(self.cb_ozel_kw.setChecked)
        # LineEdit'lerin açılıp kapanması
        self.cb_ozel_kw.toggled.connect(self.le_ozel_kw.setEnabled)
        self.cb_ozel_devir.toggled.connect(self.le_ozel_devir.setEnabled)

        duzen_ozel.addWidget(self.cb_ozel_kw,0, 0)
        duzen_ozel.addWidget(self.le_ozel_kw,0, 1)
        duzen_ozel.addWidget(self.cb_ozel_devir, 0, 2)
        duzen_ozel.addWidget(self.le_ozel_devir, 0, 3)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        # --- GRUP 4: Ek Özellikler (Onay Kutuları) ---
        grup_onaylar = QGroupBox("Ek Donanım ve Özellikler")
        duzen_onaylar = QGridLayout()

        self.cb_standart = QCheckBox("STANDART")
        self.cb_isitici = QCheckBox("ISITICI")
        self.cb_termistor = QCheckBox("TERMİSTÖR")
        self.cb_vantilator = QCheckBox("VANTİLATÖR")
        self.cb_termik = QCheckBox("TERMİK")
        self.cb_h_class = QCheckBox("H CLASS")
        self.cb_tropikalizeli = QCheckBox("TROPİKALİZELİ")

        self.cb_pt100 = QCheckBox("PT100")

        # 1. GİRİŞ GRUBU (Örn: 6X SARGIDA)
        self.pt100_adet_kutusu_1 = QComboBox()
        self.pt100_adet_kutusu_1.addItems(["", "1", "2", "3", "4", "5", "6"])
        self.pt100_adet_kutusu_1.setVisible(False)

        self.pt100_yer_kutusu_1 = QComboBox()
        self.pt100_yer_kutusu_1.addItems(["", "SARGIDA", "RULMANLARDA"])
        self.pt100_yer_kutusu_1.setVisible(False)

        # 2. GİRİŞ GRUBU (Örn: 2X RULMANDA)
        self.pt100_adet_kutusu_2 = QComboBox()
        self.pt100_adet_kutusu_2.addItems(["", "1", "2", "3", "4", "5", "6"])
        self.pt100_adet_kutusu_2.setVisible(False)

        self.pt100_yer_kutusu_2 = QComboBox()
        self.pt100_yer_kutusu_2.addItems(["", "SARGIDA", "RULMANLARDA"])
        self.pt100_yer_kutusu_2.setVisible(False)

        # Tıklanma (Check) durumuna göre TÜM 4 kutuyu göster/gizle
        self.cb_pt100.toggled.connect(self.pt100_adet_kutusu_1.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_yer_kutusu_1.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_adet_kutusu_2.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_yer_kutusu_2.setVisible)

        duzen_onaylar.addWidget(self.cb_standart, 0, 0)
        duzen_onaylar.addWidget(self.cb_isitici, 0, 1)
        duzen_onaylar.addWidget(self.cb_termistor, 0, 2)
        duzen_onaylar.addWidget(self.cb_vantilator, 0, 3)

        duzen_onaylar.addWidget(self.cb_termik, 1, 0)
        duzen_onaylar.addWidget(self.cb_h_class, 1, 1)

        duzen_onaylar.addWidget(self.cb_tropikalizeli, 1, 2)

        # 4. Satır: PT100 Checkbox ve 1. Grup
        duzen_onaylar.addWidget(self.cb_pt100, 2, 0)
        duzen_onaylar.addWidget(self.pt100_adet_kutusu_1, 2, 1)
        duzen_onaylar.addWidget(self.pt100_yer_kutusu_1, 2, 2)

        # 5. Satır: Sadece 2. Grup (0. Sütun boş, diğerlerinin altına hizalanır)
        duzen_onaylar.addWidget(self.pt100_adet_kutusu_2, 3, 1)
        duzen_onaylar.addWidget(self.pt100_yer_kutusu_2, 3, 2)

        grup_onaylar.setLayout(duzen_onaylar)
        form_ana_duzen.addWidget(grup_onaylar)

        # PT100 Kutuları İçin KESİN ÇÖZÜM: Zorla yazılabilir yapıp Placeholder ekliyoruz
        for kutu in [self.pt100_adet_kutusu_1, self.pt100_yer_kutusu_1, self.pt100_adet_kutusu_2,
                     self.pt100_yer_kutusu_2]:
            kutu.setEditable(True)

        self.pt100_adet_kutusu_1.lineEdit().setPlaceholderText("Adet/Yaz...")
        self.pt100_yer_kutusu_1.lineEdit().setPlaceholderText("1. Yer Seç/Yaz...")
        self.pt100_adet_kutusu_2.lineEdit().setPlaceholderText("Adet/Yaz...")
        self.pt100_yer_kutusu_2.lineEdit().setPlaceholderText("2. Yer Seç/Yaz...")


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
        self.tablo.setAlternatingRowColors(True)
        ana_duzen.addWidget(self.tablo, stretch=1)

        # ---------------- EXCEL'E AKTAR BUTONU ----------------
        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR (STATORLU GOVDE)")
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
            hedef_sayfa = "STATORLU GOVDE"
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.strip().upper() == "STATORLU GOVDE":
                    hedef_sayfa = sayfa
                    break

            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=None, dtype=str).fillna("")
            for idx, row in df.iterrows():
                for val in row.values:
                    val_str = str(val).strip().upper()
                    if "SG-" in val_str:
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
                if "SG-" in tablo_id:
                    parcalar = tablo_id.split("-")
                    for parca in parcalar:
                        if parca.isdigit() and len(parca) == 4:
                            kullanilan_nolar.add(int(parca))
            except:
                pass

        # 3. İLK BOŞLUĞU BUL (1'den başla, hangisi havuzda yoksa onu ver)
        for i in range(1, 10000):
            if i not in kullanilan_nolar:
                return i

        return len(kullanilan_nolar) + 1

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip()).upper()

        if not aranan_id and not aranan_aciklama:
            return

        guncel_yol = ayar_excel_yolunu_getir()
        self.tablo.setRowCount(0)

        gercek_sayfa_adi = "STATORLU GOVDE"
        try:
            excel_dosyasi = pd.ExcelFile(guncel_yol)
            for sayfa in excel_dosyasi.sheet_names:
                temiz_sayfa_adi = sayfa.strip().upper()
                if "STATORLU" in temiz_sayfa_adi and "GOVDE" in temiz_sayfa_adi:
                    gercek_sayfa_adi = sayfa
                    break
        except:
            pass

        try:
            # YENİ VE TEMİZ KISIM
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

            # YENİ: BULUNAMADI UYARISI
            if bulunan_kayit_sayisi == 0:
                QMessageBox.information(self, "Bulunamadı", "Aradığınız kriterlerde kayıt bulunamadı.")

        except ValueError:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Excel dosyasında 'STATORLU GOVDE' sayfası bulunamadı!")
        except Exception as e:
            self.tablo.setSortingEnabled(True)
            QMessageBox.warning(self, "Hata", f"Arama sırasında hata:\n{e}")

    def gelismis_arama_ac(self):
        try:
            excel_yolu = ayar_excel_yolunu_getir()
            dialog = StatorluGovdeGelismisArama(self, excel_yolu)
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Gelişmiş arama açılamadı:\n{e}")

    def excele_aktar(self):
        gercek_sayfa_adi = "STATORLU GOVDE"
        guncel_yol = ayar_excel_yolunu_getir()
        excel_yolu = os.path.abspath(guncel_yol)

        try:
            excel_dosyasi = pd.ExcelFile(excel_yolu)
            for sayfa in excel_dosyasi.sheet_names:
                temiz_sayfa = sayfa.upper().strip()
                if "STATORLU" in temiz_sayfa and "GOVDE" in temiz_sayfa:
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

            id_kolonu = None
            for col in df.columns:
                if "KOD" in col and "AÇIKLAMA" not in col:
                    id_kolonu = col
                    break

            exceldeki_idler = df[id_kolonu].astype(str).str.strip().tolist() if id_kolonu else []
            cakisan_satirlar = []
            for satir in range(self.tablo.rowCount()):
                item = self.tablo.item(satir, 0)
                if item and item.text().strip() in exceldeki_idler:
                    cakisan_satirlar.append(satir)

            if cakisan_satirlar:
                satir_numaralari = ", ".join([str(r + 1) for r in cakisan_satirlar])
                cevap_kopya = QMessageBox.question(
                    self, "KOPYA STATORLU GÖVDE UYARISI!",
                    f"Tablodaki {satir_numaralari}. satırdaki gövde(ler)in AYNISI Excel'de zaten mevcut!\n\n"
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

        # --- DİKKAT: ESKİ TOPLU_EXCELE_AKTAR FONKSİYONU SİLİNDİ, ÇÖKME ENGELLENDİ ---

        # 3. GÖLGE KOPYA (SHADOW COPY) VE YAZMA SİSTEMİ
        try:
            temp_klasor = tempfile.gettempdir()
            gecici_excel = os.path.join(temp_klasor, "gecici_aemot_statorlu_govde.xlsx")
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

            try:
                if os.path.exists(excel_yolu): os.chmod(excel_yolu, stat.S_IWRITE)
                shutil.copy2(gecici_excel, excel_yolu)
                QMessageBox.information(self, "Başarılı",
                                        "Kayıtlar Excel'in 10. satırından itibaren kusursuzca eklendi!")
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

            temiz_metin = re.sub(r'R\s*:\s*\d+', '', temiz_metin)

            for isaret in ["-", ":", "/", "\\", "_", ",", ".", "MM"]:
                temiz_metin = temiz_metin.replace(isaret, " ")
            return set(temiz_metin.split())

        programin_dizisi = kelime_dizisi_yap(taslak_metin)

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

        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = sayfa_adi
            for sayfa in excel_dosyasi.sheet_names:
                if "STATORLU" in sayfa.upper() and "GOVDE" in sayfa.upper():
                    hedef_sayfa = sayfa
                    break

            # YENİ VE TEMİZ KISIM
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]

            for index, row in df.iterrows():
                excel_grup = str(row.get('GRUP', ''))

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
            pass

        return (None, "") if resim_no_cek else None

    def uret_calistir(self):
        for kutu in self.findChildren(QComboBox):
            yeni_metin = kutu.currentText().strip().upper()
            if yeni_metin and kutu.findText(yeni_metin) == -1:
                kutu.addItem(yeni_metin)

        # 1. ARAYÜZDEN VERİLERİ ÇEK
        versiyon = self.versiyon_kutusu.currentText().strip()
        verimlilik = self.verimlilik_kutusu.currentText().strip()
        grup = self.M_kutusu.currentText().strip()
        govde_boyu = self.govde_boyu_kutusu.currentText().strip()
        kutup = self.kutup_kutusu.currentText().strip()
        paket_boyu = self.paket_boyu_kutusu.currentText().strip()
        ayak = self.ayak_kutusu.currentText().strip()
        govde_malzemesi = self.govde_malzemesi_kutusu.currentText().strip()
        resim_no_girdi = self.resim_no_kutusu.text().strip()
        revizyon = self.revizyon_kutusu.currentText().strip()
        ozel_frekans = self.ozel_frekans_kutusu.currentText().strip()
        ozel_gerilim = self.ozel_gerilim_kutusu.text().strip().upper()
        ham_ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_kutusu.text()).strip().upper()

        # --- YENİ EKLENEN: BOŞ ALAN UYARISI ---
        bos_alanlar = []
        if not grup: bos_alanlar.append("Grup")
        if not kutup: bos_alanlar.append("Kutup")
        if not govde_boyu: bos_alanlar.append("Gövde Boyu")
        if not paket_boyu: bos_alanlar.append("Paket Boyu")
        if not verimlilik: bos_alanlar.append("Verimlilik")
        if not ayak: bos_alanlar.append("Ayak")
        if not govde_malzemesi: bos_alanlar.append("Gövde Malzemesi")

        if bos_alanlar:
            cevap = QMessageBox.question(self, "Eksik Bilgi",
                                         f"Şu alanlar boş bırakıldı:\n{', '.join(bos_alanlar)}\n\nYine de üretip tabloya eklemek istiyor musunuz?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        if not resim_no_girdi:
            resim_no_girdi = "0000"

        # ONAYLAR
        standart = self.cb_standart.isChecked()
        isitici = self.cb_isitici.isChecked()
        termistor = self.cb_termistor.isChecked()
        vantilator = self.cb_vantilator.isChecked()
        termik = self.cb_termik.isChecked()
        h_class = self.cb_h_class.isChecked()
        tropikalizeli = self.cb_tropikalizeli.isChecked()
        ozel_kw = self.cb_ozel_kw.isChecked()
        kw_deger = self.le_ozel_kw.text().strip()
        devir_deger = self.le_ozel_devir.text().strip()

        # =========================================================
        # 1. KİLİT: STANDART SEÇİLİ DEĞİLSE ÖZEL KW ZORUNLUDUR!
        # =========================================================
        if not standart:
            has_kw = ozel_kw and bool(kw_deger)
            has_devir = self.cb_ozel_devir.isChecked() and bool(devir_deger)
            if not (has_kw and has_devir):
                QMessageBox.warning(self, "EKSİK BİLGİ",
                                    "Motor 'STANDART' seçilmediyse, 'Özel KW' ve 'Özel Devir' girmek zorundasınız!")
                return

        # =========================================================
        # 2. KİLİT: STANDART SEÇİLİYSE KATALOGDA VAR MI?
        # =========================================================
        if standart and not (ozel_kw and self.cb_ozel_devir.isChecked()):
            try:
                katalog = katalog_oku()
                katalog_key = f"{grup}{govde_boyu}{kutup}{paket_boyu}"

                if katalog_key not in katalog:
                    alternatifler = []
                    for k, v in katalog.items():
                        if k.startswith(grup):
                            alt_g_k_p = k[len(grup):]
                            k_w = v.get("kw", "-")
                            d_v = v.get("devir", "-")
                            alternatifler.append(f"• Gövde/Kutup/Paket: {alt_g_k_p} -> {k_w} kW / {d_v} d/dk")

                    alt_metin = "\n".join(
                        alternatifler) if alternatifler else f"Bu gruba ({grup}) ait katalogda kayıt yok."

                    cevap = QMessageBox.question(self, "STANDART KW/DEVİR BULUNAMADI!",
                                                 f"Seçtiğiniz ölçüler için ({katalog_key}) sistemde KW/Devir değeri YOK!\n\n"
                                                 f"Aynı gruba ({grup}) ait mevcut standartlar şunlardır:\n\n"
                                                 f"{alt_metin}\n\n"
                                                 f"Lütfen yukarıdaki mevcut ölçülerden birini seçin ya da 'Özel KW/Devir' kullanın.\n\n"
                                                 f"Yine de devam edip eksik KW/Devir ile kod oluşturmak istiyor musunuz?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if cevap == QMessageBox.StandardButton.No:
                        return
            except:
                pass

        yeni_eklemeler = []
        if hasattr(self, 'cb_pt100') and self.cb_pt100.isChecked():
            pt_eklendi = False
            adet_1 = self.pt100_adet_kutusu_1.currentText().strip().upper()
            yer_1 = turkce_karakter_temizleme(self.pt100_yer_kutusu_1.currentText().strip()).upper()
            adet_2 = self.pt100_adet_kutusu_2.currentText().strip().upper()
            yer_2 = turkce_karakter_temizleme(self.pt100_yer_kutusu_2.currentText().strip()).upper()

            def pt_metni_olustur(adet, yer):
                if not adet and not yer: return ""
                metin = "PT100"
                if adet: metin = f"{adet}X{metin}"
                if yer: metin = f"{yer} {metin}"
                return metin

            metin_1 = pt_metni_olustur(adet_1, yer_1)
            metin_2 = pt_metni_olustur(adet_2, yer_2)

            if metin_1:
                yeni_eklemeler.append(metin_1)
                pt_eklendi = True
            if metin_2:
                yeni_eklemeler.append(metin_2)
                pt_eklendi = True

            if not pt_eklendi:
                yeni_eklemeler.append("PT100")

        yasakli_kelimeler = ["MARIN", "MARINE", "TERMIK", "ISITICI", "GUDUK", "-V", "PASLANMAZ", "HZ",
                             "BOYA", "RAL", "IP56", "IP65", "IP66"]
        for kelime in yasakli_kelimeler:
            if kelime in ham_ek_aciklama:
                QMessageBox.critical(self, "KURAL İHLALİ",
                                     f"Bİ UYANIK SENSİN ZATEN. KOLAYA KAÇMA. KUTULARDAN SEÇEREK YAP.\n\nYasaklı Kelime: {kelime}")
                return

        # ---------------- 2. EXCEL FORMÜLÜNE GÖRE TASLAK OLUŞTURMA ----------------
        baslangic = f"{grup} {govde_boyu}{kutup}{paket_boyu}"
        if vantilator: baslangic += "-V"

        orta_1 = f" {verimlilik}" if verimlilik else ""

        # --- DİKKAT: "if not standart:" KISITLAMASI TAMAMEN KALDIRILDI! ---
        orta_2_parcalar = []
        if ozel_gerilim: orta_2_parcalar.append(f"{ozel_gerilim}V")
        if ozel_frekans and ozel_frekans != "50": orta_2_parcalar.append(f"{ozel_frekans}Hz")
        if h_class: orta_2_parcalar.append("H CLASS")

        try:
            grup_int = int(grup)
            if termistor and grup_int < 200: orta_2_parcalar.append("TERMISTORLU")
        except:
            if termistor: orta_2_parcalar.append("TERMISTORLU")

        if isitici: orta_2_parcalar.append("ISITICILI")
        if tropikalizeli: orta_2_parcalar.append("TROPIKALIZELI")
        if termik: orta_2_parcalar.append("TERMIKLI")
        if yeni_eklemeler: orta_2_parcalar.extend(yeni_eklemeler)

        orta_2 = " ".join(orta_2_parcalar)
        if orta_2: orta_2 = f" {orta_2}"

        ek_metin = f" {ham_ek_aciklama}" if ham_ek_aciklama else ""

        # Malzemeyi Excel formatına göre kısaltıyoruz
        malzeme_kisa = "AL" if govde_malzemesi == "ALUMINYUM" else govde_malzemesi
        sabit_bitis = f" {ayak} STATORLU GOVDE-{malzeme_kisa}"

        try:
            katalog = katalog_oku()
            katalog_key = f"{grup}{govde_boyu}{kutup}{paket_boyu}"
            katalog_veri = katalog.get(katalog_key, {"kw": "?", "devir": "?"})
            kw_son = kw_deger if ozel_kw else katalog_veri.get('kw', '?')
            devir_son = devir_deger if ozel_kw else katalog_veri.get('devir', '?')
        except:
            kw_son = kw_deger if ozel_kw else "?"
            devir_son = devir_deger if ozel_kw else "?"

        kw_devir_kapanis = f" - {kw_son} kW {devir_son} d/dk"
        revizyon_eki = "" if versiyon == "Eski Govde" else f" - R:{revizyon}"

        taslak_aciklama = f"{baslangic}{orta_1}{orta_2}{ek_metin}{sabit_bitis}{kw_devir_kapanis}{revizyon_eki}"
        taslak_aciklama = " ".join(taslak_aciklama.split())

        # --- YENİ EKLENDİ: SİSTEMİN AÇIKLAMA 2'Yİ OTOMATİK ÜRETMESİ VE TEMİZLİĞİ ---
        aciklama2_oto = taslak_aciklama

        if aciklama2_oto.startswith(grup):
            aciklama2_oto = aciklama2_oto[len(grup):].strip()

        if kw_devir_kapanis in aciklama2_oto:
            aciklama2_oto = aciklama2_oto.replace(kw_devir_kapanis, "").strip()
        if revizyon_eki in aciklama2_oto:
            aciklama2_oto = aciklama2_oto.replace(revizyon_eki, "").strip()

        # Sonunda gereksiz tire "-" kalırsa sil
        if aciklama2_oto.endswith("-"):
            aciklama2_oto = aciklama2_oto[:-1].strip()

        # ---------------- 3. DNA TESTİ İLE SIRA VE RESİM NO BULMA ----------------
        tam_metin_havuzu = f"{taslak_aciklama} {aciklama2_oto}"

        dna_sira, dna_resim = self.dna_bazli_sira_no_bul(sayfa_adi="STATORLU GOVDE", aranan_grup=grup,
                                                         taslak_metin=tam_metin_havuzu, resim_no_cek=True)

        resim_no = resim_no_girdi
        if (not resim_no or resim_no == "0000") and dna_resim:
            resim_no = str(dna_resim).strip()

        if resim_no.isdigit():
            resim_no = resim_no.zfill(4)
        if not resim_no or resim_no == "nan":
            resim_no = "0000"

        from arayuz.mantik.sira_no_dialog import sira_no_secim_yap
        secilen_sira = sira_no_secim_yap(self, dna_sira, taslak_aciklama, sayfa_adi="STATORLU GOVDE", id_kodu="SG")
        if secilen_sira is None:
            return
        elif secilen_sira == -1:
            sira_no_int = self.genel_max_sira_no_bul()  # DÜZELTME: fonksiyon zaten ilk boş sıra noyu döndürüyor, +1 eklemek onu dolu bir numaraya kaydırıyordu
        else:
            sira_no_int = secilen_sira

        sira_no_str = str(sira_no_int)
        aciklama_kod_str = sira_no_str.zfill(4)

        # ---------------- 4. NİHAİ ID OLUŞTURMA VE KOPYA KALKANI ----------------
        if ozel_kw:
            rev_rakam = "0" if versiyon == "Eski Govde" else revizyon[-1]
            id_revizyon = f"Q{rev_rakam}"
        else:
            id_revizyon = "00" if versiyon == "Eski Govde" else revizyon

        id_resim_no = resim_no.zfill(5) if resim_no else "0000"

        uretilen_id = f"YMM{grup}-SG-{aciklama_kod_str}-{id_revizyon}-{id_resim_no}"

        for r in range(self.tablo.rowCount()):
            if self.tablo.item(r, 0) and self.tablo.item(r, 0).text().strip() == uretilen_id:
                QMessageBox.warning(self, "KOPYA ENGELİ", f"Bu Statorlu Gövde ZATEN LİSTEDE var!\nID: {uretilen_id}")
                return

        # ---------------- 5. TABLOYA YAZDIRMA ----------------
        bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

        # Tablo Sırası: ["#KOD", "AÇIKLAMA", "GRUP", 'AÇIKLAMA2', "RESİM NO", 'AÇIKLAMA SIRA NO', 'AÇIKLAMA SIRA KOD', "KAYIT TARİHİ"]
        yeni_kayit = [
            uretilen_id, taslak_aciklama, grup, aciklama2_oto, resim_no, sira_no_str, aciklama_kod_str, bugunun_tarihi
        ]

        self.tablo.setSortingEnabled(False)
        mevcut_satir = self.tablo.rowCount()
        self.tablo.insertRow(mevcut_satir)
        for sutun, veri in enumerate(yeni_kayit):
            hucre = QTableWidgetItem(str(veri))
            self.tablo.setItem(mevcut_satir, sutun, hucre)

        self.tablo.setSortingEnabled(True)

    # ================== BAK BU FONKSİYONLAR URET_CALISTIRIN DIŞINDA OLACAK ==================
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


class StatorluGovdeGelismisArama(QDialog):
    def __init__(self, parent, excel_yolu):
        super().__init__(parent)
        self.setWindowTitle("Statorlu Gövde Gelişmiş Arama")
        self.setMinimumWidth(800)
        self.parent = parent
        self.excel_yolu = excel_yolu

        self.layout = QVBoxLayout(self)

        # --- 1. EXCEL VERİSİNİ OKU ---
        gercek_sayfa = "STATORLU GOVDE"
        try:
            dosya = pd.ExcelFile(self.excel_yolu)
            for s in dosya.sheet_names:
                temiz_sayfa = s.upper().strip()
                if "STATORLU" in temiz_sayfa and "GOVDE" in temiz_sayfa:
                    gercek_sayfa = s
                    break
            # YENİ VE TEMİZ KISIM
            self.df = pd.read_excel(self.excel_yolu, sheet_name=gercek_sayfa, header=0, dtype=str).fillna("")
            self.df.columns = [str(c).strip().upper() for c in self.df.columns]
        except:
            self.df = pd.DataFrame()

        # --- 2. FORM ALANI (STATORLU GÖVDE ANA FORMUNUN ARAMA VERSİYONU) ---
        grup_form = QGroupBox("Statorlu Gövde Parametreleriyle Filtrele")
        form_layout = QGridLayout()

        # Kutular
        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["", "Eski Govde", "Premium Govde"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["", "063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315",
             "355", "400", "450", "500", "630"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(
            ["", "2", "4", "6", "8", "10", "12", "2/4", "4/2", "4/8", "6/4", "6/8", "8/2", "8/4", "8/6", "2/12", "4/12",
             "12/4", "4/16"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(["", "S", "M", "L"])

        self.paket_boyu_kutusu = QComboBox()
        self.paket_boyu_kutusu.addItems(["", "A", "B", "C", "D", "K", "P", "X", "Y", "Z", "Q", "A-V"])

        self.ayak_kutusu = QComboBox()
        self.ayak_kutusu.addItems(["", "AYAKLI", "AYAKSIZ"])

        self.govde_malzemesi_kutusu = QComboBox()
        self.govde_malzemesi_kutusu.addItems(["", "ALUMINYUM", "PIK", "CELIK"])

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.setEditable(True)
        self.revizyon_kutusu.addItems(["", "01", "02", "03", "04", "05"])

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Ek Açıklama veya Donanım (Örn: PT100, ISITICI)")

        self.sira_no_kutusu = QLineEdit()
        self.sira_no_kutusu.setPlaceholderText("Örn: 1")

        # Yerleştirme (Grid)
        form_layout.addWidget(QLabel("Versiyon:"), 0, 0)
        form_layout.addWidget(self.versiyon_kutusu, 0, 1)
        form_layout.addWidget(QLabel("Grup (M):"), 0, 2)
        form_layout.addWidget(self.grup_kutusu, 0, 3)

        form_layout.addWidget(QLabel("Kutup Sayısı:"), 1, 0)
        form_layout.addWidget(self.kutup_kutusu, 1, 1)
        form_layout.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        form_layout.addWidget(self.govde_boyu_kutusu, 1, 3)

        form_layout.addWidget(QLabel("Paket Boyu:"), 2, 0)
        form_layout.addWidget(self.paket_boyu_kutusu, 2, 1)
        form_layout.addWidget(QLabel("Malzeme:"), 2, 2)
        form_layout.addWidget(self.govde_malzemesi_kutusu, 2, 3)

        form_layout.addWidget(QLabel("Ayak:"), 3, 0)
        form_layout.addWidget(self.ayak_kutusu, 3, 1)
        form_layout.addWidget(QLabel("Revizyon No:"), 3, 2)
        form_layout.addWidget(self.revizyon_kutusu, 3, 3)

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

            tum_metin = tum_metin.str.replace("İ", "I").str.replace("Ş", "S").str.replace("Ğ", "G").str.replace("Ü",
                                                                                                                "U").str.replace(
                "Ö", "O").str.replace("Ç", "C")
            metin_sifir_bosluk = tum_metin.str.replace(" ", "").str.replace("-", "").str.replace("/", "")

            # Form Girdilerini Çek
            versiyon = self.versiyon_kutusu.currentText().strip().upper()
            grup = self.grup_kutusu.currentText().strip().upper()
            kutup = self.kutup_kutusu.currentText().strip().upper()
            govde_boyu = self.govde_boyu_kutusu.currentText().strip().upper()
            paket_boyu = self.paket_boyu_kutusu.currentText().strip().upper()
            ayak = self.ayak_kutusu.currentText().strip().upper()
            malzeme = self.govde_malzemesi_kutusu.currentText().strip().upper()
            revizyon = self.revizyon_kutusu.currentText().strip().upper()
            ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
            sira_no = self.sira_no_kutusu.text().strip()

            # --- AKILLI FİLTRELEME KURALLARI ---
            if grup and len(df_filtrelenmis.columns) > 2:
                df_filtrelenmis = df_filtrelenmis[
                    df_filtrelenmis.iloc[:, 2].astype(str).str.strip().str.upper() == grup]

            if kutup:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"-{kutup}", regex=False) | tum_metin.str.contains(f" {kutup} ",
                                                                                              regex=False)]

            if govde_boyu:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(govde_boyu, regex=False)]

            if paket_boyu:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(paket_boyu, regex=False)]

            if ayak:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(ayak, regex=False)]

            if malzeme:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(malzeme, regex=False)]

            if revizyon:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"R:{revizyon}", regex=False) | tum_metin.str.contains(f"-{revizyon}-",
                                                                                                  regex=False)]

            if ek_aciklama:
                ea_temiz = ek_aciklama.replace(" ", "")
                df_filtrelenmis = df_filtrelenmis[metin_sifir_bosluk.str.contains(ea_temiz, regex=False)]

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