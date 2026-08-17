from PyQt6.QtWidgets import (
    QWidget, QComboBox, QGridLayout, QLineEdit, QSpacerItem, QSizePolicy,
    QLabel, QCheckBox, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QPushButton, QDialog, QVBoxLayout,
    QHBoxLayout, QFormLayout, QCompleter, QGroupBox, QScrollArea
)
from PyQt6.QtGui import QPixmap, QIcon, QFont
from PyQt6.QtCore import QTimer, QDateTime, Qt
import traceback
import pandas as pd
import json
import os
import tempfile, shutil, stat, os
import xlwings as xw
from arayuz.mantik.motor_verileri import motor, turkce_karakter_temizleme, motor_ozellikleri
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.json_islemleri import katalog_oku, standart_flans_oku, ayar_excel_yolunu_getir, kaynak_yolu
from arayuz.mantik.kopya_bulucu import ayni_motor_var_mi
from arayuz.mantik import oturum

class MotorArayuzu(QWidget):
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

        # ANA DÜZEN
        # ANA DÜZEN
        ana_duzen = QVBoxLayout()
        self.setLayout(ana_duzen)

        # ---------------- 1. ÜST BUTONLAR VE BİLGİ ALANI (HEADER) ----------------
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

        header_duzen = QHBoxLayout()

        self.logo_etiketi = QLabel()
        self.logo_etiketi.setPixmap(QPixmap(kaynak_yolu("arayuz/varliklar/aemot_logo_gorseli.png")))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)

        baslik = QLabel("MOTOR İD VE AÇIKLAMA OLUŞTURUCU")
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
        grup_temel = QGroupBox("Temel Motor Özellikleri")
        duzen_temel = QGridLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["Eski Govde", "Premium Govde", "PM Govde"])
        self.verimlilik_kutusu = QComboBox()
        self.verimlilik_kutusu.addItems(["IE1", "IE2", "IE3", "IE4", "IE5"])
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
        self.paket_boyu_kutusu.addItems(["A", "B", "C", "D", "E", "F", "K", "P", "T", "X", "Y", "Q", "RH", "A-V"])
        self.yapi_sekli_kutusu = QComboBox()
        self.yapi_sekli_kutusu.addItems(["B3", "B5", "B9", "B14", "B34","B35"])
        self.flans_olcusu_kutusu = QComboBox()
        self.flans_olcusu_kutusu.addItems(
            ["YOK", "C90", "C105", "C120", "C140", "C160", "C200", "C250", "A140", "A160", "A200", "A250", "A300",
             "A350", "A400", "A450", "A550", "A660", "A800", "A1000", "A1150", "A1370"])

        duzen_temel.addWidget(QLabel("Versiyon:"), 0, 0)
        duzen_temel.addWidget(self.versiyon_kutusu, 0, 1)
        duzen_temel.addWidget(QLabel("Verimlilik:"), 0, 2)
        duzen_temel.addWidget(self.verimlilik_kutusu, 0, 3)
        duzen_temel.addWidget(QLabel("M (Gövde):"), 1, 0)
        duzen_temel.addWidget(self.M_kutusu, 1, 1)
        duzen_temel.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        duzen_temel.addWidget(self.govde_boyu_kutusu, 1, 3)

        duzen_temel.addWidget(QLabel("Kutup:"), 2, 0)
        duzen_temel.addWidget(self.kutup_kutusu, 2, 1)
        duzen_temel.addWidget(QLabel("Paket Boyu:"), 2, 2)
        duzen_temel.addWidget(self.paket_boyu_kutusu, 2, 3)
        duzen_temel.addWidget(QLabel("Yapı Şekli:"), 3, 0)
        duzen_temel.addWidget(self.yapi_sekli_kutusu, 3, 1)
        duzen_temel.addWidget(QLabel("Flanş Ölçüsü:"), 3, 2)
        duzen_temel.addWidget(self.flans_olcusu_kutusu, 3, 3)

        grup_temel.setLayout(duzen_temel)
        form_ana_duzen.addWidget(grup_temel)

        # --- GRUP 2: Opsiyonel ve Serbest Girişler ---
        grup_opsiyon = QGroupBox("Opsiyonel Özellikler")
        duzen_opsiyon = QGridLayout()

        self.govde_malzemesi_kutusu = QComboBox()
        self.govde_malzemesi_kutusu.addItems(["GOVDESIZ", "ALUMINYUM", "PIK", "CELIK"])

        self.resim_no_input = QLineEdit()
        self.resim_no_input.setPlaceholderText("Örn: 3130 (Boş bırakılabilir)")

        self.ek_aciklama_kutusu = QLineEdit()
        self.ek_aciklama_kutusu.setPlaceholderText("Ek açıklama giriniz...")

        self.ip_sinifi_kutusu = QComboBox()
        self.ip_sinifi_kutusu.addItems(["55", "56", "65", "66", "67"])
        self.ozel_frekans_kutusu = QComboBox()
        self.ozel_frekans_kutusu.addItems(["50","33",  "60", "67", "75", "87", "100", "106", "120", "133", "135.33"])
        self.ozel_gerilim_kutusu = QLineEdit()
        self.ozel_gerilim_kutusu.setPlaceholderText("Varsa özel gerilim")
        self.paslanmaz_civata_kutusu = QComboBox()
        self.paslanmaz_civata_kutusu.addItems(["", "INOX", "A-4/316"])

        duzen_opsiyon.addWidget(QLabel("Gövde Malzemesi:"), 0, 0)
        duzen_opsiyon.addWidget(self.govde_malzemesi_kutusu, 0, 1)
        duzen_opsiyon.addWidget(QLabel("Resim No:"), 0, 2)
        duzen_opsiyon.addWidget(self.resim_no_input, 0, 3)
        duzen_opsiyon.addWidget(QLabel("Ek Açıklama:"), 1, 0)
        duzen_opsiyon.addWidget(self.ek_aciklama_kutusu, 1, 1, 1, 3)

        duzen_opsiyon.addWidget(QLabel("IP Sınıfı:"), 2, 0)
        duzen_opsiyon.addWidget(self.ip_sinifi_kutusu, 2, 1)
        duzen_opsiyon.addWidget(QLabel("Özel Frekans:"), 2, 2)
        duzen_opsiyon.addWidget(self.ozel_frekans_kutusu, 2, 3)
        duzen_opsiyon.addWidget(QLabel("Özel Gerilim:"), 3, 0)
        duzen_opsiyon.addWidget(self.ozel_gerilim_kutusu, 3, 1)
        duzen_opsiyon.addWidget(QLabel("Paslanmaz Cıvata:"), 3, 2)
        duzen_opsiyon.addWidget(self.paslanmaz_civata_kutusu, 3, 3)

        grup_opsiyon.setLayout(duzen_opsiyon)
        form_ana_duzen.addWidget(grup_opsiyon)

        # --- GRUP 3: Boya, Özel KW ve Özel Devir ---
        grup_ozel = QGroupBox("Boya ve Özel Değerler")
        duzen_ozel = QGridLayout()

        self.cb_boya = QCheckBox("Boya Durumu")
        self.le_boya = QLineEdit()
        self.le_boya.setPlaceholderText("Örn: 3020 EPOXY")
        self.le_boya.setEnabled(False)
        self.cb_boya.toggled.connect(self.le_boya.setEnabled)

        self.cb_ozel_kw = QCheckBox("Özel KW")
        self.le_ozel_kw = QLineEdit()
        self.le_ozel_kw.setPlaceholderText("Örn: 1.5")
        self.le_ozel_kw.setEnabled(False)

        self.cb_ozel_devir = QCheckBox("Özel Devir")
        self.le_ozel_devir = QLineEdit()
        self.le_ozel_devir.setPlaceholderText("Örn: 1800")
        self.le_ozel_devir.setEnabled(False)

        self.cb_ozel_kw.toggled.connect(self.cb_ozel_devir.setChecked)
        self.cb_ozel_devir.toggled.connect(self.cb_ozel_kw.setChecked)
        self.cb_ozel_kw.toggled.connect(self.le_ozel_kw.setEnabled)
        self.cb_ozel_devir.toggled.connect(self.le_ozel_devir.setEnabled)

        duzen_ozel.addWidget(self.cb_boya, 0, 0)
        duzen_ozel.addWidget(self.le_boya, 0, 1)
        duzen_ozel.addWidget(self.cb_ozel_kw, 0, 2)
        duzen_ozel.addWidget(self.le_ozel_kw, 0, 3)
        duzen_ozel.addWidget(self.cb_ozel_devir, 0, 4)
        duzen_ozel.addWidget(self.le_ozel_devir, 0, 5)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        # --- GRUP 4: Ek Donanım ve Özellikler ---
        grup_onaylar = QGroupBox("Ek Donanım ve Özellikler")
        duzen_onaylar = QGridLayout()

        self.cb_standart = QCheckBox("STANDART")
        self.cb_isitici = QCheckBox("ISITICI")
        self.cb_marine = QCheckBox("MARINE")
        self.cb_guduk = QCheckBox("GÜDÜK")

        self.cb_termistor = QCheckBox("TERMİSTÖR")
        self.cb_vantilator = QCheckBox("VANTİLATÖR")
        self.cb_termik = QCheckBox("TERMİK")
        self.cb_paslanmaz_rakor = QCheckBox("PASLANMAZ RAKOR")

        self.cb_tropikalizeli = QCheckBox("TROPİKALİZELİ")
        self.cb_cd = QCheckBox("ÇEKTİRME DELİKSİZ")  # CD'nin ismini güncelledik
        self.cb_onden_sabit = QCheckBox("ÖNDEN SABİT")
        self.cb_enmotor = QCheckBox("EN MOTOR")

        self.cb_imak = QCheckBox("İMAK")
        self.cb_regl = QCheckBox("REGAL")
        self.cb_teco = QCheckBox("TECO")
        self.izole_rulmanda = QCheckBox("IZOLE RULMANLI")
        self.cb_nu_rulman = QCheckBox("NU RULMANLI")
        self.cb_h_class = QCheckBox("H CLASS")
        self.cb_corona_telli = QCheckBox("CORONA TELLİ")

        # YENİ: ÇAKIŞMA ENGELLEYİCİLER (MUTEX KİLİTLERİ)
        # NU RULMAN ve ÖNDEN SABİT çakışması
        self.cb_nu_rulman.toggled.connect(lambda checked: self.cb_onden_sabit.setEnabled(not checked))
        self.cb_onden_sabit.toggled.connect(lambda checked: self.cb_nu_rulman.setEnabled(not checked))

        # MARINE ve PASLANMAZ çakışması
        def marine_kontrol(checked):
            self.cb_paslanmaz_rakor.setEnabled(not checked)
            self.paslanmaz_civata_kutusu.setEnabled(not checked)
            if checked:
                self.cb_paslanmaz_rakor.setChecked(False)
                self.paslanmaz_civata_kutusu.setCurrentIndex(0)

        def paslanmaz_kontrol():
            rakor_dolu = self.cb_paslanmaz_rakor.isChecked()
            civata_dolu = bool(self.paslanmaz_civata_kutusu.currentText().strip())
            self.cb_marine.setEnabled(not (rakor_dolu or civata_dolu))
            if rakor_dolu or civata_dolu:
                self.cb_marine.setChecked(False)

        self.cb_marine.toggled.connect(marine_kontrol)
        self.cb_paslanmaz_rakor.toggled.connect(lambda _: paslanmaz_kontrol())
        self.paslanmaz_civata_kutusu.currentTextChanged.connect(lambda _: paslanmaz_kontrol())

        # YENİ: GİZLENEBİLİR VE 4 KUTULU (ÇİFT GİRİŞLİ) PT100 MENÜSÜ
        self.cb_pt100 = QCheckBox("PT100")

        # 1. GİRİŞ GRUBU (Örn: 6X SARGIDA)
        self.pt100_adet_kutusu_1 = QComboBox()
        self.pt100_adet_kutusu_1.addItems(["", "1", "2", "3", "4", "5", "6"])
        self.pt100_adet_kutusu_1.setVisible(False)

        self.pt100_yer_kutusu_1 = QComboBox()
        self.pt100_yer_kutusu_1.addItems(["", "SARGIDA"])
        self.pt100_yer_kutusu_1.setVisible(False)

        # 2. GİRİŞ GRUBU (Örn: 2X RULMANDA)
        self.pt100_adet_kutusu_2 = QComboBox()
        self.pt100_adet_kutusu_2.addItems(["", "1", "2", "3", "4", "5", "6"])
        self.pt100_adet_kutusu_2.setVisible(False)

        self.pt100_yer_kutusu_2 = QComboBox()
        self.pt100_yer_kutusu_2.addItems(["","RULMANDA"])
        self.pt100_yer_kutusu_2.setVisible(False)

        # Tıklanma (Check) durumuna göre TÜM 4 kutuyu göster/gizle
        self.cb_pt100.toggled.connect(self.pt100_adet_kutusu_1.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_yer_kutusu_1.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_adet_kutusu_2.setVisible)
        self.cb_pt100.toggled.connect(self.pt100_yer_kutusu_2.setVisible)

        # --- EKRANA YERLEŞTİRME KISMI (Grid Matrisi) ---
        duzen_onaylar.addWidget(self.cb_standart, 0, 0)
        duzen_onaylar.addWidget(self.cb_isitici, 0, 1)
        duzen_onaylar.addWidget(self.cb_marine, 0, 2)
        duzen_onaylar.addWidget(self.cb_guduk, 0, 3)

        duzen_onaylar.addWidget(self.cb_termistor, 1, 0)
        duzen_onaylar.addWidget(self.cb_vantilator, 1, 1)
        duzen_onaylar.addWidget(self.cb_termik, 1, 2)
        duzen_onaylar.addWidget(self.cb_paslanmaz_rakor, 1, 3)

        duzen_onaylar.addWidget(self.cb_tropikalizeli, 2, 0)
        duzen_onaylar.addWidget(self.cb_cd, 2, 1)  # Çektirme Deliksiz
        duzen_onaylar.addWidget(self.cb_onden_sabit, 2, 2)  # Önden Sabit
        duzen_onaylar.addWidget(self.cb_nu_rulman, 2, 3)  # İŞTE BURAYA OTURTTUK (NU Rulman)

        duzen_onaylar.addWidget(self.cb_imak, 3, 0)
        duzen_onaylar.addWidget(self.cb_regl, 3, 1)
        duzen_onaylar.addWidget(self.cb_teco, 3, 2)
        duzen_onaylar.addWidget(self.izole_rulmanda, 3, 3)

        duzen_onaylar.addWidget(self.cb_enmotor, 4, 0)
        duzen_onaylar.addWidget(self.cb_h_class, 4, 1)
        duzen_onaylar.addWidget(self.cb_corona_telli, 4, 2)
        # 4, 1 ve 4, 2 boş kalabilir, sıkıntı yok

        # 4. ve 5. Satır: PT100 Checkbox ve Girişleri
        duzen_onaylar.addWidget(self.cb_pt100, 5, 0)
        duzen_onaylar.addWidget(self.pt100_adet_kutusu_1, 5, 1)
        duzen_onaylar.addWidget(self.pt100_yer_kutusu_1, 5, 2, 1, 2)

        duzen_onaylar.addWidget(self.pt100_adet_kutusu_2, 6, 1)
        duzen_onaylar.addWidget(self.pt100_yer_kutusu_2, 6, 2, 1, 2)

        grup_onaylar.setLayout(duzen_onaylar)
        form_ana_duzen.addWidget(grup_onaylar)

        # --- BÜTÜN COMBOBOX'LARI MANUEL YAZIMA AÇMA ---
        for kutu in self.findChildren(QComboBox):
            kutu.setEditable(True)
            kutu.lineEdit().setPlaceholderText("Seç veya Yaz...")

            if kutu.findText("") == -1:
                kutu.insertItem(0, "")

            kutu.lineEdit().mousePressEvent = lambda event, le=kutu.lineEdit(): le.selectAll()
            kutu.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            completer = kutu.completer()
            if completer:
                completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                completer.setFilterMode(Qt.MatchFlag.MatchContains)

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

        # YENİDEN EKLENEN ANA MENÜ BUTONU

        self.uret_butonu = QPushButton("LİSTEYE EKLE (ID ÜRET)")
        self.uret_butonu.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.uret_butonu.clicked.connect(self.kayit_baslat)

        buton_duzeni.addWidget(self.uret_butonu, stretch=2)

        ana_duzen.addLayout(buton_duzeni)

        # ---------------- 4. ARAMA VE TABLO ALANI (TEMİZLENDİ) ----------------
        arama_duzeni = QHBoxLayout()

        self.id_arama_kutusu = QLineEdit()
        self.id_arama_kutusu.setPlaceholderText("ID Ara (Örn: 355M4A)...")

        self.arama_kutusu = QLineEdit()
        self.arama_kutusu.setPlaceholderText("Açıklamada Ara (Örn: 3020 EPOXY)...")

        self.btn_ara = QPushButton("Ara")
        self.btn_ara.setStyleSheet("background-color: #3498DB; color: white; font-weight: bold; padding: 5px;")
        self.btn_ara.clicked.connect(self.tabloda_arama_yap)

        self.btn_gelismis_ara = QPushButton("Gelişmiş Arama")
        self.btn_gelismis_ara.setStyleSheet("background-color: #E67E22; color: white; font-weight: bold; padding: 5px;")
        self.btn_gelismis_ara.clicked.connect(self.gelismis_arama_ac)

        self.btn_satir_sil = QPushButton("Seçili Satırı Sil")
        self.btn_satir_sil.setStyleSheet("background-color: #C0392B; color: white; font-weight: bold; padding: 5px;")
        self.btn_satir_sil.clicked.connect(self.secili_satiri_sil)

        self.btn_tam_ekran = QPushButton("🗖")
        self.btn_tam_ekran.setFixedSize(32, 32)
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

        # Tablo Oluşturma
        self.tablo = QTableWidget(0, 16)
        self.tablo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.tablo_basliklari = [
            "#KOD", "AÇIKLAMA", "VERİMLİLİK", "GRUP", "GÖVDE BOYU",
            "KUTUP", "PAKET BOYU", "YAPI ŞEKLİ", "FLANŞ ÖLÇÜSÜ",
            "GÖVDE MALZEMESİ", "RESİM NO", "EK AÇIKLAMA",
            "AÇIKLAMA NO", "AÇIKLAMA KOD", "KW/DEVİR", "KAYIT TARİHİ"
        ]
        self.tablo.setHorizontalHeaderLabels(self.tablo_basliklari)
        self.tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tablo.setAlternatingRowColors(True)

        # Filtreleme / Sıralama
        self.tablo.setSortingEnabled(True)

        # Sağ tık ve kopyalama özellikleri
        self.tablo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tablo.customContextMenuRequested.connect(self.tablo_sag_tik)
        self.tablo.keyPressEvent = self.tablo_kopyalama_dinleyicisi

        ana_duzen.addWidget(self.tablo, stretch=1)

        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR")
        self.aktar_butonu.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.aktar_butonu.clicked.connect(self.excele_aktar_arayuz)
        ana_duzen.addWidget(self.aktar_butonu)

        self.zamanlayici = QTimer(self)
        self.zamanlayici.timeout.connect(self.tarih_mevcut_gosterim)
        self.zamanlayici.start(1000)

    # =========================================================================
    # FONKSİYONLAR
    # =========================================================================

    def yeni_sira_no_bul(self):
        kullanilan_nolar = set()

        # 1. EXCEL'DEKİ TÜM AÇIKLAMA SIRA NOLARI TOPLA
        try:
            excel_verisi = pd.read_excel(ayar_excel_yolunu_getir(), dtype=str, header=0).fillna("")
            temiz_kolonlar = {turkce_karakter_temizleme(str(col)).strip().upper(): col for col in excel_verisi.columns}

            # Arama yaparken ASLA Türkçe karakter kullanmıyoruz
            sira_kolonu = temiz_kolonlar.get('ACIKLAMA NO', temiz_kolonlar.get('ACIKLAMA KOD', temiz_kolonlar.get('ACIKLAMA SIRA NO')))

            if sira_kolonu:
                for index, row in excel_verisi.iterrows():
                    val = str(row.get(sira_kolonu, '0')).replace("'", "").strip().split(".")[0]
                    val = val.replace("Q", "").replace("q", "")
                    if val and val.isdigit():
                        no = int(val)
                        if no > 0: kullanilan_nolar.add(no)
        except Exception as e:
            print("Sıra No Okuma Hatası:", e)

        # 2. ARAYÜZ TABLOSUNDAKİ (Henüz Kaydedilmemiş) NUMARALARI TOPLA
        for row in range(self.tablo.rowCount()):
            try:
                # Ana Motor Tablosunda 12. indeks AÇIKLAMA NO sütunudur
                val = self.tablo.item(row, 12).text().strip().split(".")[0]
                if val and val.isdigit():
                    no = int(val)
                    if no > 0: kullanilan_nolar.add(no)
            except:
                pass

        # 3. İLK BOŞLUĞU BUL (Rotorlu Mildeki gibi aradaki boşlukları doldurur)
        for i in range(1, 100000):
            if i not in kullanilan_nolar:
                return i
        return 1

    def dna_bazli_sira_no_bul(self, aranan_grup, taslak_metin):
        def kelime_kumesi_yap(metin):
            temiz = turkce_karakter_temizleme(str(metin)).upper()
            import re
            temiz = re.sub(r'R\s*:\s*\d+', '', temiz)  # Revizyon atlama kuralı
            for isaret in ["-", ":", "/", "\\", "_", ",", ".", "MM", "KW", "D", "DK", "RPM"]:
                temiz = temiz.replace(isaret, " ")
            return set([k.strip() for k in temiz.split() if k.strip()])

        programin_dna_kumesi = kelime_kumesi_yap(taslak_metin)

        # Eğer ek açıklama bomboşsa (Standart motorsa)
        if not programin_dna_kumesi:
            return {0: {'adet': 1, 'tam_aciklama': 'STANDART MOTOR', 'id': 'STANDART'}}

        # Sözlük Yapısı: {1: {'adet': 5, 'tam_aciklama': '...', 'id': '...'}, 2: {...}}
        bulunanlar_sozlugu = {}

        # 1. Arayüz Tablosunda Ara
        for row_idx in range(self.tablo.rowCount()):
            try:
                tab_ek_aciklama = str(self.tablo.item(row_idx, 11).text())
                if kelime_kumesi_yap(tab_ek_aciklama) == programin_dna_kumesi:
                    no_str = self.tablo.item(row_idx, 12).text().strip().split(".")[0]
                    no_str = no_str.replace("Q", "").replace("q", "")
                    if no_str.isdigit():
                        no = int(no_str)
                        if no not in bulunanlar_sozlugu:
                            bulunanlar_sozlugu[no] = {
                                'adet': 0,
                                'tam_aciklama': self.tablo.item(row_idx, 1).text(),  # Ana Açıklama
                                'id': self.tablo.item(row_idx, 0).text()  # Kod
                            }
                        bulunanlar_sozlugu[no]['adet'] += 1
            except:
                pass

        # 2. Excel Geçmişinde Ara
        try:
            excel_verisi = pd.read_excel(ayar_excel_yolunu_getir(), dtype=str, header=0).fillna("")
            temiz_excel_kolonlari = {turkce_karakter_temizleme(str(col)).strip().upper(): col for col in
                                     excel_verisi.columns}

            sira_kolonu_adi = temiz_excel_kolonlari.get('ACIKLAMA NO', temiz_excel_kolonlari.get('ACIKLAMA KOD',
                                                                                                 temiz_excel_kolonlari.get(
                                                                                                     'ACIKLAMA SIRA NO')))
            ek_aciklama_kolonu = temiz_excel_kolonlari.get('EK ACIKLAMA')
            ana_aciklama_kolonu = temiz_excel_kolonlari.get('ACIKLAMA')
            kod_kolonu = temiz_excel_kolonlari.get('KOD', temiz_excel_kolonlari.get('#KOD'))

            if sira_kolonu_adi and ek_aciklama_kolonu:
                for index, row in excel_verisi.iterrows():
                    excel_ek_aciklama = str(row.get(ek_aciklama_kolonu, ''))

                    if kelime_kumesi_yap(excel_ek_aciklama) == programin_dna_kumesi:
                        val = str(row.get(sira_kolonu_adi, '0')).replace("'", "").strip().split(".")[0]
                        val = val.replace("Q", "").replace("q", "")
                        if val and val.isdigit():
                            no = int(val)
                            if no not in bulunanlar_sozlugu:
                                # Bu sıra noyu ilk defa görüyorsak örnek olarak bu kaydı al
                                bulunanlar_sozlugu[no] = {
                                    'adet': 0,
                                    'tam_aciklama': str(row.get(ana_aciklama_kolonu, '')),
                                    'id': str(row.get(kod_kolonu, ''))
                                }
                            # Her bulduğunda sadece o sıra nonun adetini 1 arttır
                            bulunanlar_sozlugu[no]['adet'] += 1
        except Exception as e:
            print("DNA Okuma Hatası:", e)

        return bulunanlar_sozlugu

    def sira_no_secim_paneli_goster(self, bulunanlar_sozlugu, aranan_ek_aciklama):
        dialog = QDialog(self)
        dialog.setWindowTitle("Sıra Numarası Çakışması Tespit Edildi!")
        dialog.setMinimumSize(850, 400)
        layout = QVBoxLayout(dialog)

        bilgi_etiketi = QLabel(
            f"Aşağıdaki EK AÇIKLAMA için geçmişte farklı sıra numaraları kullanılmış:\n\n"
            f"'{aranan_ek_aciklama}'\n\n"
            f"Lütfen listeye eklemek istediğiniz Sıra Numarasını seçin veya Yeni Numara Üretin."
        )
        bilgi_etiketi.setWordWrap(True)
        bilgi_etiketi.setStyleSheet("font-weight: bold; color: #C0392B; font-size: 14px;")
        layout.addWidget(bilgi_etiketi)

        # Tablo Profili
        tablo = QTableWidget(len(bulunanlar_sozlugu), 4)
        tablo.setHorizontalHeaderLabels(
            ["Sıra No", "Kullanım Adedi", "Örnek Motor Kodu", "Ana Motor Açıklaması (Örnek)"])
        tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        tablo.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        tablo.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tablo.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        tablo.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        tablo.setAlternatingRowColors(True)

        satir_index = 0
        siralari_sirala = sorted(bulunanlar_sozlugu.keys())
        for no in siralari_sirala:
            detay = bulunanlar_sozlugu[no]

            item_no = QTableWidgetItem(str(no).zfill(4))
            item_no.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_no.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))

            item_adet = QTableWidgetItem(f"{detay['adet']} Adet")
            item_adet.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_adet.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            item_adet.setForeground(Qt.GlobalColor.darkGreen)

            item_id = QTableWidgetItem(detay['id'])
            item_id.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            item_aciklama = QTableWidgetItem(detay['tam_aciklama'])

            tablo.setItem(satir_index, 0, item_no)
            tablo.setItem(satir_index, 1, item_adet)
            tablo.setItem(satir_index, 2, item_id)
            tablo.setItem(satir_index, 3, item_aciklama)
            satir_index += 1

        tablo.selectRow(0)
        layout.addWidget(tablo)

        # ---------------- BUTONLAR ----------------
        buton_duzeni = QHBoxLayout()
        btn_sec = QPushButton("Seçili Sıra Numarasını Uygula")
        btn_sec.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")

        btn_yeni_uret = QPushButton("Yeni Sıra No Üret")
        btn_yeni_uret.setStyleSheet(
            "background-color: #3498DB; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")

        btn_iptal = QPushButton("Kayıt İşlemini İptal Et")
        btn_iptal.setStyleSheet(
            "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")

        btn_elle_gir = QPushButton("🔒 Elle Sıra No Gir")
        btn_elle_gir.setStyleSheet(
            "background-color: #8E44AD; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        buton_duzeni.addWidget(btn_elle_gir)

        buton_duzeni.addWidget(btn_sec)
        buton_duzeni.addWidget(btn_yeni_uret)
        buton_duzeni.addWidget(btn_iptal)
        layout.addLayout(buton_duzeni)

        secilen_no = [None]

        def on_sec_clicked():
            secili_satirlar = tablo.selectionModel().selectedRows()
            if secili_satirlar:
                secilen_no[0] = int(tablo.item(secili_satirlar[0].row(), 0).text())
                dialog.accept()
            else:
                QMessageBox.warning(dialog, "Uyarı", "Lütfen tablodan bir numara seçin!")

        def on_yeni_uret_clicked():
            secilen_no[0] = -1  # Sisteme yepyeni numara üretmesi için verdiğimiz gizli sinyal
            dialog.accept()

        def on_elle_gir_clicked():
            from PyQt6.QtWidgets import QInputDialog, QLineEdit
            sifre, ok = QInputDialog.getText(dialog, "Yetkili Girişi", "Lütfen yetkili şifresini girin:",
                                             QLineEdit.EchoMode.Password)
            if ok and sifre == "aemot123":
                no, ok2 = QInputDialog.getInt(dialog, "Elle Sıra No Gir", "Atamak istediğiniz sıra numarasını girin:",
                                              min=1, max=99999)
                if ok2:
                    secilen_no[0] = no
                    dialog.accept()
            elif ok:
                QMessageBox.warning(dialog, "Yetkisiz Erişim", "Hatalı şifre girdiniz!")

        btn_sec.clicked.connect(on_sec_clicked)
        btn_yeni_uret.clicked.connect(on_yeni_uret_clicked)
        btn_iptal.clicked.connect(dialog.reject)
        btn_elle_gir.clicked.connect(on_elle_gir_clicked)

        sonuc = dialog.exec()
        if sonuc == QDialog.DialogCode.Accepted:
            return secilen_no[0]
        return None


    def kayit_baslat(self):
        for kutu in self.findChildren(QComboBox):
            yeni_metin = kutu.currentText().strip().upper()
            if yeni_metin and kutu.findText(yeni_metin) == -1:
                kutu.addItem(yeni_metin)
        try:
            # 1. BOŞ ALAN UYARISI VE DEĞİŞKENLERİ EN TEPEDEN SABİTLEME
            grup_val = self.M_kutusu.currentText().strip()
            kutup_val = self.kutup_kutusu.currentText().strip()
            govde_val = self.govde_boyu_kutusu.currentText().strip()
            paket_val = self.paket_boyu_kutusu.currentText().strip()
            verimlilik_val = self.verimlilik_kutusu.currentText().strip()
            yapi_val = self.yapi_sekli_kutusu.currentText().strip().upper()
            flans_val = self.flans_olcusu_kutusu.currentText().strip().upper()
            versiyon_val = self.versiyon_kutusu.currentText().strip()

            bos_alanlar = []
            if not grup_val: bos_alanlar.append("Grup")
            if not kutup_val: bos_alanlar.append("Kutup")
            if not govde_val: bos_alanlar.append("Gövde Boyu")
            if not paket_val: bos_alanlar.append("Paket Boyu")
            if not yapi_val: bos_alanlar.append("Yapı Şekli")

            if bos_alanlar:
                cevap = QMessageBox.question(self, "Eksik Bilgi",
                                             f"Şu alanlar boş:\n{', '.join(bos_alanlar)}\n\nYine de üretip tabloya eklemek istiyor musunuz?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if cevap == QMessageBox.StandardButton.No:
                    return

            ham_aciklama = self.ek_aciklama_kutusu.text()
            temiz_aciklama = turkce_karakter_temizleme(ham_aciklama).replace("  ", " ").strip()

            if not self.cb_standart.isChecked():
                has_kw = self.cb_ozel_kw.isChecked() and bool(self.le_ozel_kw.text().strip())
                has_devir = self.cb_ozel_devir.isChecked() and bool(self.le_ozel_devir.text().strip())
                if not (has_kw and has_devir):
                    QMessageBox.warning(self, "EKSİK BİLGİ",
                                        "Motor 'STANDART' seçilmediyse, 'Özel KW' ve 'Özel Devir' girmek zorundasınız!")
                    return

            # ---------------------------------------------------------
            # STANDART KW/DEVİR KONTROLÜ VE ÖZEL UYARI PANELİ
            # ---------------------------------------------------------
            if self.cb_standart.isChecked() and not (self.cb_ozel_kw.isChecked() and self.cb_ozel_devir.isChecked()):
                try:
                    katalog = katalog_oku()
                    katalog_key = f"{grup_val}{govde_val}{kutup_val}{paket_val}"

                    if katalog_key not in katalog:
                        alternatifler = []
                        for k, v in katalog.items():
                            if k.startswith(grup_val):
                                alt_g_k_p = k[len(grup_val):]
                                k_w = v.get("kw", "-")
                                d_v = v.get("devir", "-")
                                alternatifler.append(f"• Gövde/Kutup/Paket: {alt_g_k_p} -> {k_w} kW / {d_v} d/dk")

                        alt_metin = "\n".join(
                            alternatifler) if alternatifler else f"Bu gruba ({grup_val}) ait katalogda kayıt yok."

                        cevap = QMessageBox.question(self, "STANDART KW/DEVİR BULUNAMADI!",
                                                     f"Seçtiğiniz ölçüler için ({katalog_key}) sistemde KW/Devir değeri YOK!\n\n"
                                                     f"Aynı gruba ({grup_val}) ait mevcut standartlar şunlardır:\n\n"
                                                     f"{alt_metin}\n\n"
                                                     f"Lütfen yukarıdaki mevcut ölçülerden birini seçin ya da 'Özel KW/Devir' kullanın.\n\n"
                                                     f"Yine de devam edip eksik KW/Devir ile kod oluşturmak istiyor musunuz?",
                                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                        if cevap == QMessageBox.StandardButton.No:
                            return
                except:
                    pass

            # ---------------------------------------------------------
            # NOKTA ATIŞI STANDART FLANŞ KONTROLÜ VE ONAY
            # ---------------------------------------------------------
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

                        if m_json.lstrip("0") == grup_val.lstrip("0"):
                            if yapi_json in ["B5", "B14", "B34","B35"]:
                                alternatifler.append(f"• Yapı Şekli: {yapi_json} -> Flanş: {flans_json}")
                            if yapi_json == yapi_val:
                                beklenen_flans = flans_json
            except Exception as e:
                pass

            if self.cb_standart.isChecked() and yapi_val in ["B5", "B14", "B34","B35"]:
                if not beklenen_flans:
                    alt_metin = "\n".join(sorted(list(
                        set(alternatifler)))) if alternatifler else f"Bu gruba ({grup_val}) ait tanımlı standart flanş kaydı yok."
                    cevap = QMessageBox.question(self, "STANDART FLANŞ BULUNAMADI!",
                                                 f"Seçtiğiniz Gövde ({grup_val}) ve Yapı Şekli ({yapi_val}) için sistemde tanımlı bir flanş değeri YOK!\n\n"
                                                 f"Aynı gruba ({grup_val}) ait mevcut standartlar şunlardır:\n\n{alt_metin}\n\n"
                                                 f"Yine de flanş ölçüsünü '{flans_val}' olarak ayarlayıp kod üretmek istiyor musunuz?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if cevap == QMessageBox.StandardButton.No: return

                elif beklenen_flans and flans_val != beklenen_flans:
                    cevap = QMessageBox.question(self, "STANDART FLANŞ UYUMSUZLUĞU!",
                                                 f"STANDART motor seçtiniz ancak girdiğiniz flanş ölçüsü hatalı!\n\n"
                                                 f"Gövde (M): {grup_val}\nYapı Şekli: {yapi_val}\nSizin Seçiminiz: {flans_val if flans_val and flans_val != 'YOK' else 'GİRİLMEDİ (YOK)'}\n\n"
                                                 f"Olması Gereken Standart Flanş: {beklenen_flans}\n\n"
                                                 f"Yine de flanşı '{flans_val}' olarak kabul edip üretime devam etmek istiyor musunuz?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if cevap == QMessageBox.StandardButton.No: return

            uyarilar = []
            if yapi_val in ["B3", "B9"] and flans_val != "YOK":
                uyarilar.append(f"{yapi_val} yapı şeklinde Flanş 'YOK' olmalıdır.")
            elif yapi_val not in ["B3", "B9"]:
                if flans_val == "YOK":
                    flans_mesaji = f"{yapi_val} yapı şeklinde Flanş GİRİLMELİDİR."
                    if beklenen_flans: flans_mesaji += f" (Olması gereken standart: {beklenen_flans})"
                    uyarilar.append(flans_mesaji)
                elif beklenen_flans and flans_val != beklenen_flans:
                    uyarilar.append(
                        f"Girdiğiniz flanş ({flans_val}) standart dışıdır! (Olması gereken standart: {beklenen_flans})")

            try:
                m_degeri = int(grup_val)
            except:
                m_degeri = 0

            malzeme = self.govde_malzemesi_kutusu.currentText().strip().upper()

            if 63 <= m_degeri <= 112 and malzeme != "ALUMINYUM":
                uyarilar.append(f"{m_degeri} gövde standart ALÜMİNYUM olmalıdır.")
            elif 225 <= m_degeri <= 400 and malzeme != "PIK":
                uyarilar.append(f"{m_degeri} gövde standart PIK olmalıdır.")
            elif m_degeri > 400 and malzeme != "CELIK":
                uyarilar.append(f"{m_degeri} gövde standart ÇELİK olmalıdır.")
            if m_degeri >= 200 and malzeme == "ALUMINYUM": uyarilar.append(f"{m_degeri} gövdede ALÜMİNYUM üretilmez!")

            if uyarilar:
                uyari_metni = "\n".join([f"• {u}" for u in uyarilar])
                cevap = QMessageBox.question(self, "STANDART DIŞI ONAYI",
                                             f"Standart dışı seçimler tespit edildi:\n\n{uyari_metni}\n\nYine de kod üretmek istiyor musunuz?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if cevap == QMessageBox.StandardButton.No: return

            # YASAKLI KELİMELER KONTROLÜ
            yasakli_kelimeler = ["TROPIKALIZELI", "MARIN", "MARINE", "TERMISTOR", "TERMIK", "ISITICI", "GUDUK", "-V",
                                 "PASLANMAZ", "HZ", "RAL", "IP56", "IP65", "IP66", "IP67", "PT100", "NU RULMAN",
                                 "CEKTIRME"]
            for kelime in yasakli_kelimeler:
                if kelime in temiz_aciklama.upper():
                    QMessageBox.critical(self, "KURAL İHLALİ",
                                         f"Yasaklı Kelime: {kelime}\nBunu ek açıklama yerine onay kutularından seçiniz.")
                    return

            ip_degeri_kutu = self.ip_sinifi_kutusu.currentText().strip()
            # Değişkeni doldur ama kutu boşsa varsayılan olarak kabul et
            ip_degeri = ip_degeri_kutu or "55"
            marine_secili = self.cb_marine.isChecked()
            termistor_secili = self.cb_termistor.isChecked()

            if marine_secili and ip_degeri == "56": ip_degeri = ""
            if marine_secili: termistor_secili = False
            if m_degeri >= 200: termistor_secili = False

            # YENİ: RESİM NO BOŞ İSE "0000" YAPMA
            resim_no_girdi = self.resim_no_input.text().strip()
            if not resim_no_girdi:
                resim_no_girdi = "0000"

            # ---------------------------------------------------------
            # YENİ: KUSURSUZ SIRALAMAYA GÖRE EK AÇIKLAMA ÜRETİMİ
            # SIRA: Resim No > Özel Gerilim > Özel Frekans > Boya > IP > MARINE > GUDUK > ISITICILI > TROPIKALIZELI > TERMIKLI > PASLANMAZ > PT100 > RULMANLAR > EK_ACIKLAMA > MARKALAR
            # ---------------------------------------------------------
            ek_donanimlar = []

            # 1. Resim No
            if resim_no_girdi != "0000":
                ek_donanimlar.append(resim_no_girdi)

            # 2. Özel Gerilim
            ozel_gerilim = turkce_karakter_temizleme(self.ozel_gerilim_kutusu.text().strip()).upper()
            if ozel_gerilim:
                # Kullanıcı sonuna V koymayı unuttuysa otomatik ekle
                if not ozel_gerilim.endswith("V"):
                    ozel_gerilim += "V"
                ek_donanimlar.append(ozel_gerilim)

            # 3. Özel Frekans
            ozel_frekans = self.ozel_frekans_kutusu.currentText().strip() or "50"
            # 3. Özel Frekans (50 ise EK AÇIKLAMAYA YAZMA)
            if ozel_frekans and ozel_frekans != "50":
                ek_donanimlar.append(f"{ozel_frekans}Hz")


            # 4. Boya (RAL ve BOYASIZ KONTROLÜ)
            temiz_boya = turkce_karakter_temizleme(self.le_boya.text().strip()).upper()
            if self.cb_boya.isChecked():
                if not temiz_boya or "BOYASIZ" in temiz_boya:
                    ek_donanimlar.append("BOYASIZ")
                else:
                    boya_metni = temiz_boya
                    if "RAL" not in boya_metni:
                        boya_metni = f"RAL {boya_metni}"
                    if "BOYALI" not in boya_metni:
                        boya_metni = f"{boya_metni} BOYALI"
                    ek_donanimlar.append(boya_metni)

            # 5. IP (Kutuda bir şey yazıyorsa ve 55 DEĞİLSE açıklamaya ekle)
            if ip_degeri_kutu and ip_degeri_kutu != "55":
                ek_donanimlar.append(f"IP{ip_degeri_kutu}")

            # 6. Marine ve Guduk
            if self.cb_vantilator.isChecked(): ek_donanimlar.append("-V")
            if marine_secili: ek_donanimlar.append("MARINE")
            if self.cb_guduk.isChecked(): ek_donanimlar.append("GUDUK")
            if self.cb_isitici.isChecked(): ek_donanimlar.append("ISITICILI")
            if self.cb_tropikalizeli.isChecked(): ek_donanimlar.append("TROPIKALIZELI")
            if self.cb_termik.isChecked(): ek_donanimlar.append("TERMIKLI")
            if termistor_secili: ek_donanimlar.append("TERMISTORLU")

            # 7. Paslanmaz Rakor ve Civata
            if self.cb_paslanmaz_rakor.isChecked(): ek_donanimlar.append("PASLANMAZ RAKORLU")
            paslanmaz_civata = turkce_karakter_temizleme(self.paslanmaz_civata_kutusu.currentText().strip()).upper()
            if paslanmaz_civata: ek_donanimlar.append(f"{paslanmaz_civata} CIVATALI")

            if hasattr(self, 'cb_h_class') and self.cb_h_class.isChecked(): ek_donanimlar.append("H CLASS")
            if hasattr(self, 'cb_corona_telli') and self.cb_corona_telli.isChecked(): ek_donanimlar.append(
                "CORONA TELLI")
            # 8. PT100 Matrisi
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
                    ek_donanimlar.append(metin_1)
                    pt_eklendi = True
                if metin_2:
                    ek_donanimlar.append(metin_2)
                    pt_eklendi = True
                if not pt_eklendi:
                    ek_donanimlar.append("PT100")

            # 9. Rulmanlar ve Sabitlemeler
            if self.izole_rulmanda.isChecked(): ek_donanimlar.append("IZOLE RULMANLI")
            if hasattr(self, 'cb_nu_rulman') and self.cb_nu_rulman.isChecked(): ek_donanimlar.append("NU RULMANLI")
            if self.cb_onden_sabit.isChecked(): ek_donanimlar.append("ONDEN SABIT")

            if hasattr(self, 'cb_cd') and self.cb_cd.isChecked():
                ek_donanimlar.append("CEKTIRME DELIKSIZ")

            # 10. MANUEL YAZILAN EK AÇIKLAMA
            if temiz_aciklama: ek_donanimlar.append(temiz_aciklama)

            # 11. Markalar ve Motor Tipleri
            if self.cb_enmotor.isChecked(): ek_donanimlar.append("EN MOTOR")
            if self.cb_imak.isChecked(): ek_donanimlar.append("IMAK")
            if self.cb_regl.isChecked(): ek_donanimlar.append("REGAL")
            if self.cb_teco.isChecked(): ek_donanimlar.append("TECO")

            # TABLOYA VE DNA'YA GİDECEK KUSURSUZ EK AÇIKLAMA METNİ
            tablo_icin_ek_aciklama = " ".join(ek_donanimlar).strip()

            # ---------------------------------------------------------
            # YENİ: KUSURSUZ ANA AÇIKLAMANIN OLUŞTURULMASI
            # ---------------------------------------------------------
            # Format: "IE1 090 S2A-V B5 A200 5555 230/666 75Hz RAL 3020 BOYALI IP66 MARINE GUDUK ISITICILI TROPIKALIZELI TERMIKLI PASLANMAZ RAKORLU SARGIDA 6XPT100 RULMANLARDA 2XPT100 IZOLE RULMANLI NU RULMANLI ONDEN SABIT CEKTIRME DELIKSIZ EKACIKLAMA EN MOTOR IMAK REGAL TECO ALUMINYUM GOVDELI MOTOR - 1,5 kW 1800 d/dk - PREMIUM"

            vantilator_eki = "-V" if self.cb_vantilator.isChecked() else ""
            uretilen_aciklama = f"{verimlilik_val} {grup_val} {govde_val}{kutup_val}{paket_val}{vantilator_eki} {yapi_val}"

            if flans_val != "YOK":
                uretilen_aciklama += f" {flans_val}"

            if tablo_icin_ek_aciklama:
                uretilen_aciklama += f" {tablo_icin_ek_aciklama}"

            # KW ve DEVİR Hesabı
            kw, devir = "-", "-"
            if self.cb_ozel_kw.isChecked() and self.cb_ozel_devir.isChecked():
                kw = self.le_ozel_kw.text().strip()
                devir = self.le_ozel_devir.text().strip()
            else:
                try:
                    katalog_key = f"{grup_val}{govde_val}{kutup_val}{paket_val}"
                    veri = katalog_oku().get(katalog_key, {"kw": "-", "devir": "-"})
                    kw, devir = veri.get('kw', '-'), veri.get('devir', '-')
                except:
                    pass

            kw_devir_metni = f"{kw} kW / {devir} d/dk"

            # Açıklamanın Kuyruk Kısmı (Malzeme ve kW/Devir)
            # PM Gövde Kuralı (Malzemeden önce MIKNATISLI ekle)
            malzeme_metni = malzeme
            if versiyon_val == "PM Govde":
                # Eğer zaten ek açıklamalarda MIKNATISLI yoksa ekle
                if "MIKNATISLI" not in uretilen_aciklama:
                    malzeme_metni = f"MIKNATISLI {malzeme}"

            # Açıklamanın Kuyruk Kısmı
            uretilen_aciklama += f" {malzeme_metni} GOVDELI MOTOR - {kw} kW {devir} d/dk"

            if versiyon_val == "Premium Govde":
                uretilen_aciklama += " - PREMIUM"
            # KOPYA KONTROLÜ (Aynı Özellikler Zaten Excel'de Var mı?)
            kontrol_ozellikleri = {
                'VERİMLİLİK': verimlilik_val, 'GRUP': grup_val, 'GÖVDE BOYU': govde_val, 'KUTUP': kutup_val,
                'PAKET BOYU': paket_val, 'YAPI ŞEKLİ': yapi_val, 'FLANŞ ÖLÇÜSÜ': flans_val, 'GÖVDE MALZEMESİ': malzeme,
                'EK AÇIKLAMA': tablo_icin_ek_aciklama,'KW/DEVİR': kw_devir_metni
            }


            if not hasattr(self, 'excel_verisi') or self.excel_verisi is None:
                try:
                    self.excel_verisi = pd.read_excel(ayar_excel_yolunu_getir(), dtype=str, header=0).fillna("")
                except:
                    pass

            if hasattr(self, 'excel_verisi') and self.excel_verisi is not None:
                kopya_var_mi, kopya_satir = ayni_motor_var_mi(self.excel_verisi, kontrol_ozellikleri)
                if kopya_var_mi:
                    kopya_id = kopya_satir.get('KOD', kopya_satir.get('#KOD', 'Bulunamadı'))
                    cevap = QMessageBox.question(self, "BU MOTOR ZATEN VAR!",
                                                 f"Girdiğiniz özelliklerin BİREBİR AYNISI Excel'de kayıtlı.\n\nID: {kopya_id}\n\nYine de bu motoru tabloda görmek ister misiniz?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if cevap == QMessageBox.StandardButton.Yes:
                        self.tablo.setSortingEnabled(False)
                        mevcut_satir = self.tablo.rowCount()
                        self.tablo.insertRow(mevcut_satir)
                        kopya_veri = [
                            kopya_satir.get('KOD', kopya_satir.get('#KOD', '')), kopya_satir.get('AÇIKLAMA', ''),
                            kopya_satir.get('VERİMLİLİK', ''), kopya_satir.get('GRUP', ''),
                            kopya_satir.get('GÖVDE BOYU', ''),
                            kopya_satir.get('KUTUP', ''), kopya_satir.get('PAKET BOYU', ''),
                            kopya_satir.get('YAPI ŞEKLİ', ''),
                            kopya_satir.get('FLANŞ ÖLÇÜSÜ', ''), kopya_satir.get('GÖVDE MALZEMESİ', ''),
                            kopya_satir.get('RESİM NO', ''),
                            kopya_satir.get('EK AÇIKLAMA', ''), kopya_satir.get('AÇIKLAMA NO', ''),
                            kopya_satir.get('AÇIKLAMA KOD', ''),
                            kopya_satir.get('KW/DEVİR', ''), QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")
                        ]
                        for sutun, veri in enumerate(kopya_veri):
                            self.tablo.setItem(mevcut_satir, sutun, QTableWidgetItem(str(veri)))
                        self.tablo.setSortingEnabled(True)
                    return

            # SIRA NO VE ID ÜRETİMİ
            # SIRA NO VE ID ÜRETİMİ
            bulunanlar_sozlugu = self.dna_bazli_sira_no_bul(grup_val, tablo_icin_ek_aciklama)

            if bulunanlar_sozlugu:
                if len(bulunanlar_sozlugu) == 1:
                    # Tek bir sıra no var (Standart 0 veya tek benzersiz numara)
                    aciklama_no_int = list(bulunanlar_sozlugu.keys())[0]
                else:
                    # FARKLI NUMARALAR VARSA: Kullanıcıya tabloyu göster
                    secilen_deger = self.sira_no_secim_paneli_goster(bulunanlar_sozlugu, tablo_icin_ek_aciklama)

                    if secilen_deger is None:
                        # Kullanıcı iptal etti
                        return
                    elif secilen_deger == -1:
                        # Kullanıcı "Yeni Sıra No Üret" butonuna bastı
                        aciklama_no_int = self.yeni_sira_no_bul()
                    else:
                        # Kullanıcı tablodan var olan bir numarayı seçti
                        aciklama_no_int = secilen_deger
            else:
                # Normal şartlarda buraya düşmez ama garanti olsun
                aciklama_no_int = self.yeni_sira_no_bul()

            aciklama_no_str = str(aciklama_no_int)
            if self.cb_ozel_kw.isChecked() or self.cb_ozel_devir.isChecked():
                aciklama_kod_str = f"Q{aciklama_no_str.zfill(3)}"
            else:
                aciklama_kod_str = aciklama_no_str.zfill(4)

            # Sadece tip ID oluşturmak için sınıfı çağırıyoruz (açıklama parametresini boş gönderiyoruz)
            # Sadece tip ID oluşturmak için sınıfı çağırıyoruz (tüm zorunlu parametreleri vererek)
            olusturulan_motor = motor(
                versiyon=versiyon_val, verimlilik=verimlilik_val, M=grup_val, govde_boyu=govde_val,
                kutup=kutup_val, paket_boyu=paket_val, yapi_sekli=yapi_val, flans_olcusu=flans_val,
                govde_malzemesi=malzeme, resim_no=resim_no_girdi, sira_no=aciklama_kod_str, ek_aciklama="",
                ip_sinifi=ip_degeri, ozel_frekans=ozel_frekans,
                ozel_gerilim=ozel_gerilim,
                paslanmaz_civata=paslanmaz_civata,
                boya_durum=self.cb_boya.isChecked(), boya_deger=temiz_boya,
                ozel_kw_durum=self.cb_ozel_kw.isChecked(), ozel_kw_deger=self.le_ozel_kw.text().strip(),
                ozel_devir_durum=self.cb_ozel_devir.isChecked(), ozel_devir_deger=self.le_ozel_devir.text().strip(),
                standart=self.cb_standart.isChecked(),
                isitici=self.cb_isitici.isChecked(), marine=marine_secili, guduk=self.cb_guduk.isChecked(),
                termistor=termistor_secili, vantilator=self.cb_vantilator.isChecked(),
                termik=self.cb_termik.isChecked(),
                paslanmaz_rakor=self.cb_paslanmaz_rakor.isChecked(), tropikalizeli=self.cb_tropikalizeli.isChecked()
            )
            uretilen_id = olusturulan_motor.tip_id_uret()
            uretilen_id = olusturulan_motor.tip_id_uret()

            # --- KOPYA KALKANI (HEM ID HEM AÇIKLAMA BİREBİR AYNI OLMAK ZORUNDA) ---
            for r in range(self.tablo.rowCount()):
                tablo_id = self.tablo.item(r, 0).text().strip() if self.tablo.item(r, 0) else ""
                tablo_aciklama = self.tablo.item(r, 1).text().strip() if self.tablo.item(r, 1) else ""

                if tablo_id == uretilen_id and tablo_aciklama == uretilen_aciklama:
                    QMessageBox.warning(self, "KOPYA ENGELİ",
                                        "Bu motor ZATEN LİSTEDE var!\nAynı motoru ard arda ekleyemezsiniz.")
                    return

            bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

            yeni_kayit = [
                uretilen_id, uretilen_aciklama, verimlilik_val, grup_val, govde_val,
                kutup_val, paket_val, yapi_val, flans_val,
                malzeme, resim_no_girdi, tablo_icin_ek_aciklama, aciklama_no_str, aciklama_kod_str, kw_devir_metni,
                bugunun_tarihi
            ]

            self.tablo.setSortingEnabled(False)
            mevcut_satir = self.tablo.rowCount()
            self.tablo.insertRow(mevcut_satir)
            for sutun, veri in enumerate(yeni_kayit):
                self.tablo.setItem(mevcut_satir, sutun, QTableWidgetItem(str(veri)))

            self.tablo.setSortingEnabled(True)

        except Exception as e:
            hata_mesaji = traceback.format_exc()
            QMessageBox.critical(self, "SİSTEM ÇÖKMEKTEN KURTARILDI",
                                 f"Kayıt sırasında bir hata oluştu:\n\n{hata_mesaji}")

    def tarih_mevcut_gosterim(self):
        self.saat_etiketi.setText(QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm:ss"))

    def excele_aktar_arayuz(self):


        try:
            if self.tablo.rowCount() == 0:
                QMessageBox.warning(self, "Uyarı", "Tabloda aktarılacak veri yok!")
                return

            guncel_yol = ayar_excel_yolunu_getir()
            excel_yolu = os.path.abspath(guncel_yol)

            # 1. EXCEL'İ OKU VE SAYFAYI BUL
            try:
                excel_dosyasi = pd.ExcelFile(excel_yolu)
                # Genelde ana motor ilk sayfadır, garantilemek için ilkini alıyoruz
                gercek_sayfa_adi = excel_dosyasi.sheet_names[0]
                self.excel_verisi = pd.read_excel(excel_yolu, sheet_name=gercek_sayfa_adi, dtype=str, header=0).fillna(
                    "")
            except Exception as e:
                QMessageBox.warning(self, "Hata", f"Excel okunamadı! Yolu kontrol edin.\n{e}")
                return

            # 2. KOPYA KONTROLÜ (DNA BAZLI VE DİNAMİK)
            df_temiz = self.excel_verisi.copy()
            df_temiz.columns = [str(c).strip().upper() for c in df_temiz.columns]

            aciklama_idx = None
            for idx, col in enumerate(df_temiz.columns):
                if ("AÇIKLAMA" in col or "ACIKLAMA" in col) and "2" not in col and "NO" not in col and "KOD" not in col:
                    aciklama_idx = idx
                    break

            excel_aciklama_dnalari = []
            if aciklama_idx is not None:
                for deger in df_temiz.iloc[:, aciklama_idx].astype(str).tolist():
                    dna = turkce_karakter_temizleme(deger).upper().replace(" ", "")
                    if dna and dna != "NAN":
                        excel_aciklama_dnalari.append(dna)

            cakisan_satirlar = []
            for satir in range(self.tablo.rowCount()):
                item = self.tablo.item(satir, 1)  # Tabloda 1. İndeks AÇIKLAMA sütunudur
                if item:
                    tablo_dna = turkce_karakter_temizleme(item.text().strip()).upper().replace(" ", "")
                    if tablo_dna and tablo_dna in excel_aciklama_dnalari:
                        cakisan_satirlar.append(satir)

            if cakisan_satirlar:
                satir_numaralari = ", ".join([str(r + 1) for r in cakisan_satirlar])
                cevap = QMessageBox.question(
                    self, "KOPYA MOTOR UYARISI!",
                    f"Tablodaki {satir_numaralari}. satırdaki motor(lar)ın AÇIKLAMASI Excel'de zaten mevcut!\n\n"
                    f"• EVET: Kopyaları tablodan kaldır ve sadece YENİLERİ kaydet.\n"
                    f"• İPTAL: İşlemi durdur.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel
                )
                if cevap == QMessageBox.StandardButton.Cancel:
                    return
                elif cevap == QMessageBox.StandardButton.Yes:
                    for r in sorted(cakisan_satirlar, reverse=True): self.tablo.removeRow(r)
                    if self.tablo.rowCount() == 0:
                        QMessageBox.information(self, "Bilgi", "Kopyalar silinince tabloda kaydedilecek veri kalmadı.")
                        return

            # 3. YENİLERİ AKTARIM ONAYI
            excel_aktarim_cevap = QMessageBox.question(self, "KAYIT ONAYI",
                                                       f"Tablodaki verileri Excel'e ({gercek_sayfa_adi}) GÖNDERMEK istiyor musunuz?",
                                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if excel_aktarim_cevap == QMessageBox.StandardButton.No: return

            # 4. GÖLGE KOPYA VE KULLANICI ZIMBALAMA
            temp_klasor = tempfile.gettempdir()
            gecici_excel = os.path.join(temp_klasor, "gecici_aemot_anamotor.xlsx")
            shutil.copy2(excel_yolu, gecici_excel)
            os.chmod(gecici_excel, stat.S_IWRITE)

            app = xw.App(visible=False)
            app.display_alerts = False

            wb = app.books.open(gecici_excel)
            sheet = wb.sheets[gercek_sayfa_adi]

            tablo_sutun_sayisi = self.tablo.columnCount()
            import string
            harfler = list(string.ascii_uppercase) + [f"A{h}" for h in string.ascii_uppercase]
            son_sutun_harfi = harfler[tablo_sutun_sayisi]  # Örn: P'den sonraki harf (Q)

            # Başlıkları A1 hizasına yazıyoruz
            sheet.range(f'{harfler[tablo_sutun_sayisi - 1]}1').value = "KAYIT TARİHİ"
            sheet.range(f'{son_sutun_harfi}1').value = "MOTORU EKLEYEN KİŞİ"
            try:
                sheet.range('A1:Z1').api.Font.Bold = True
            except:
                pass

            son_satir = sheet.range('A1048576').end('up').row
            if son_satir < 1: son_satir = 1

            aktarilacak_veriler = []
            for r in range(self.tablo.rowCount()):
                satir_verisi = []
                for c in range(tablo_sutun_sayisi):
                    item = self.tablo.item(r, c)
                    satir_verisi.append(item.text() if item else "")

                # SATIR SONUNA KULLANICIYI EKLE
                satir_verisi.append(oturum.tam_isim())
                aktarilacak_veriler.append(satir_verisi)

            if aktarilacak_veriler:
                baslama_satiri = son_satir + 1
                bitis_satiri = baslama_satiri + len(aktarilacak_veriler) - 1
                hedef_alan = sheet.range(f'A{baslama_satiri}:{son_sutun_harfi}{bitis_satiri}')
                hedef_alan.number_format = '@'
                hedef_alan.value = aktarilacak_veriler

            wb.save(gecici_excel)
            wb.close()
            app.quit()

            try:
                if os.path.exists(excel_yolu): os.chmod(excel_yolu, stat.S_IWRITE)
                shutil.copy2(gecici_excel, excel_yolu)
                QMessageBox.information(self, "Başarılı", "Kayıtlar Excel'e kusursuzca eklendi!")
                self.tablo.setRowCount(0)
            except PermissionError:
                QMessageBox.critical(self, "DOSYA KİLİTLİ!", "Excel dosyası şu an ağda başka biri tarafından AÇIK!")
            except Exception as e:
                QMessageBox.critical(self, "Aktarım Hatası", f"Dosya yerine konurken hata oluştu:\n{e}")

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            QMessageBox.critical(self, "SİSTEM ÇÖKMEKTEN KURTARILDI",
                                 f"Excel'e aktarırken hata:\n\n{traceback.format_exc()}")

    def formu_temizle(self):
        for nesne in self.findChildren(QComboBox):
            nesne.setCurrentIndex(0)
        for nesne in self.findChildren(QLineEdit):
            nesne.clear() # Tüm QLineEdit'leri (Arama kutuları dahil) temizler
        for nesne in self.findChildren(QCheckBox):
            nesne.setChecked(False)
        self.tablo.setRowCount(0)
        self.satir_sayaci.setText("Bulunan Kayıt: 0")

    def secili_satiri_sil(self):
        secili_satirlar = self.tablo.selectionModel().selectedRows()
        if not secili_satirlar:
            QMessageBox.warning(self, "Uyarı", "Lütfen silmek için tablodan bir satır seçin!")
            return
        for satir in sorted(secili_satirlar, reverse=True): self.tablo.removeRow(satir.row())

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip()).upper()

        if not aranan_id and not aranan_aciklama: return
        self.tablo.setRowCount(0)

        try:
            excel_verisi = pd.read_excel(ayar_excel_yolunu_getir(), dtype=str, header=0).fillna("")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Excel okunamadı! Hata: {e}")
            return

        temiz_excel_kolonlari = {turkce_karakter_temizleme(str(col)).strip().upper(): col for col in
                                 excel_verisi.columns}

        id_kolonu = None
        for col in temiz_excel_kolonlari.keys():
            if "KOD" in col and "AÇIKLAMA" not in col and "ACIKLAMA" not in col:
                id_kolonu = temiz_excel_kolonlari[col]
                break

        self.tablo.setSortingEnabled(False)
        bulunan = 0

        for index, row in excel_verisi.iterrows():
            id_eslesti = True
            aciklama_eslesti = True

            if aranan_id:
                satir_id = turkce_karakter_temizleme(str(row[id_kolonu]).strip()).upper() if id_kolonu else ""
                if aranan_id not in satir_id: id_eslesti = False

            if aranan_aciklama:
                # DİKKAT: 'AÇIKLAMA' yerine 'ACIKLAMA' arıyoruz çünkü kolon isimlerindeki Türkçe karakterleri sildik!
                kolon_aciklama = temiz_excel_kolonlari.get('ACIKLAMA', '')
                kolon_ek_aciklama = temiz_excel_kolonlari.get('EK ACIKLAMA', '')

                veri_aciklama = str(row[kolon_aciklama]) if kolon_aciklama in row else ""
                veri_ek_aciklama = str(row[kolon_ek_aciklama]) if kolon_ek_aciklama in row else ""

                tam_aciklama = f"{veri_aciklama} {veri_ek_aciklama}"
                tam_aciklama_havuzu = turkce_karakter_temizleme(tam_aciklama).upper()

                aranan_kelimeler = aranan_aciklama.split()
                if not all(kelime in tam_aciklama_havuzu for kelime in aranan_kelimeler):
                    aciklama_eslesti = False

            if id_eslesti and aciklama_eslesti:
                mevcut_satir = self.tablo.rowCount()
                self.tablo.insertRow(mevcut_satir)

                if id_kolonu and pd.notna(row[id_kolonu]):
                    self.tablo.setItem(mevcut_satir, 0, QTableWidgetItem(str(row[id_kolonu]).strip()))

                # DİNAMİK KOLON EŞLEŞTİRME
                for c in range(self.tablo.columnCount()):
                    tablo_baslik_adi = turkce_karakter_temizleme(self.tablo.horizontalHeaderItem(c).text()).upper()
                    if tablo_baslik_adi in temiz_excel_kolonlari:
                        gercek_kolon = temiz_excel_kolonlari[tablo_baslik_adi]
                        veri = row[gercek_kolon]
                        if pd.notna(veri):
                            self.tablo.setItem(mevcut_satir, c, QTableWidgetItem(str(veri).strip()))
                bulunan += 1

        self.tablo.setSortingEnabled(True)
        self.satir_sayaci.setText(f"Bulunan Kayıt: {bulunan}")
        if bulunan == 0:
            QMessageBox.information(self, "Bulunamadı", "Aradığınız kriterlere uygun kayıt bulunamadı.")


    def tam_ekran_modu_degistir(self):
        su_an_gorunur_mu = self.form_alani_widgeti.isVisible()
        self.form_alani_widgeti.setVisible(not su_an_gorunur_mu)
        if su_an_gorunur_mu:
            self.btn_tam_ekran.setText("Formu Geri Getir")
            self.btn_tam_ekran.setStyleSheet(
                "background-color: #27AE60; color: white; font-weight: bold; padding: 5px;")
        else:
            self.btn_tam_ekran.setText("Tabloyu Büyük Ekran Yap")
            self.btn_tam_ekran.setStyleSheet(
                "background-color: #8E44AD; color: white; font-weight: bold; padding: 5px;")

    def tablo_sag_tik(self, pos):
        from PyQt6.QtWidgets import QMenu
        menu = QMenu()
        kopyala_aksiyon = menu.addAction("Seçili Alanı Kopyala (Ctrl+C)")
        secim = menu.exec(self.tablo.viewport().mapToGlobal(pos))
        if secim == kopyala_aksiyon: self.kopyala_islev()

    def tablo_kopyalama_dinleyicisi(self, event):
        from PyQt6.QtWidgets import QApplication
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

    def gelismis_arama_ac(self):
        try:
            excel_yolu = ayar_excel_yolunu_getir()
            dialog = MotorGelismisArama(self, excel_yolu)
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Gelişmiş arama açılamadı:\n{e}")


# ---------------- GELİŞMİŞ ARAMA SINIFI ----------------
class MotorGelismisArama(QDialog):
    def __init__(self, parent, excel_yolu):
        super().__init__(parent)
        self.setWindowTitle("Gelişmiş Arama Paneli")
        self.setMinimumWidth(550)
        self.parent = parent
        self.excel_yolu = excel_yolu

        from PyQt6.QtWidgets import QDateEdit, QLineEdit
        from PyQt6.QtCore import QDate

        self.layout = QVBoxLayout(self)
        form_layout = QGridLayout()

        self.arama_kutu_sozlugu = {}

        try:
            self.df = pd.read_excel(self.excel_yolu, dtype=str, header=0).fillna("")
            self.temiz_kolonlar = {turkce_karakter_temizleme(str(col)).strip().upper(): col for col in self.df.columns}
        except:
            self.df = pd.DataFrame()
            self.temiz_kolonlar = {}

        secenekler = {
            "GRUP": ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315",
                     "355", "400", "450", "500", "630"],
            "KUTUP": ["2", "4", "6", "8", "10", "12", "2/4", "4/2", "4/8", "6/4", "6/8", "8/2", "8/4", "8/6", "2/12",
                      "4/12", "12/4", "4/16"],
            "GÖVDE BOYU": ["S", "M", "L"],
            "YAPI ŞEKLİ": ["B3", "B5", "B9", "B14", "B34","B35"],
            "FLANŞ ÖLÇÜSÜ": ["YOK", "C90", "C105", "C120", "C140", "C160", "C200", "C250", "A140", "A160", "A200",
                             "A250", "A300", "A350", "A400", "A450", "A550", "A660", "A800", "A1000", "A1150", "A1370"],
            "VERİMLİLİK": ["IE1", "IE2", "IE3", "IE4", "IE5"],
            "RESİM NO": [],
            "EK AÇIKLAMA": []
        }

        sira = [
            ("GRUP", 0, 0), ("KUTUP", 0, 2),
            ("GÖVDE BOYU", 1, 0), ("YAPI ŞEKLİ", 1, 2),
            ("FLANŞ ÖLÇÜSÜ", 2, 0), ("VERİMLİLİK", 2, 2),
            ("RESİM NO", 3, 0), ("EK AÇIKLAMA", 3, 2)
        ]

        for baslik, satir, sutun in sira:
            kutu = QComboBox()
            kutu.setEditable(True)
            kutu.lineEdit().setPlaceholderText(f"{baslik} seç veya yaz...")
            kutu.lineEdit().mousePressEvent = lambda event, le=kutu.lineEdit(): le.selectAll()

            kutu.addItem("")
            if secenekler[baslik]:
                kutu.addItems(secenekler[baslik])
            else:
                temiz_baslik = turkce_karakter_temizleme(baslik).upper()
                if temiz_baslik in self.temiz_kolonlar and not self.df.empty:
                    gercek_kolon = self.temiz_kolonlar[temiz_baslik]
                    benzersiz_veriler = self.df[gercek_kolon].unique().tolist()
                    temiz_veriler = sorted([str(x).strip() for x in benzersiz_veriler if str(x).strip()])
                    kutu.addItems(temiz_veriler)

            form_layout.addWidget(QLabel(baslik + ":"), satir, sutun)
            form_layout.addWidget(kutu, satir, sutun + 1)
            self.arama_kutu_sozlugu[turkce_karakter_temizleme(baslik).upper()] = kutu

        # --- YENİ: SIRA NO (KOD) KUTUSU ---
        self.sira_no_kutusu = QLineEdit()
        self.sira_no_kutusu.setPlaceholderText("Örn: 1 (0001 olarak arar)")
        form_layout.addWidget(QLabel("Sıra No (Kod):"), 4, 0)
        form_layout.addWidget(self.sira_no_kutusu, 4, 1)

        # --- Kayıt Tarihi Aralığı ---
        grup_tarih = QGroupBox("Kayıt Tarihi Aralığı")
        duzen_tarih = QHBoxLayout()

        self.cb_tarih_kullan = QCheckBox("Tarihe Göre Ara")
        self.cb_tarih_kullan.setChecked(False)

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

        self.layout.addLayout(form_layout)

        # --- YENİ: CHECKBOX FİLTRELERİ ---
        grup_checkboxlar = QGroupBox("Ek Donanım ve Özellik Filtreleri")
        duzen_cb = QGridLayout()

        self.cb_filtre_guduk = QCheckBox("GÜDÜK")
        self.cb_filtre_marine = QCheckBox("MARINE")
        self.cb_filtre_cd = QCheckBox("ÇEKTİRME DELİKSİZ")
        self.cb_filtre_pt100 = QCheckBox("PT100")
        self.cb_filtre_isitici = QCheckBox("ISITICI")
        self.cb_filtre_imak = QCheckBox("İMAK")

        duzen_cb.addWidget(self.cb_filtre_guduk, 0, 0)
        duzen_cb.addWidget(self.cb_filtre_marine, 0, 1)
        duzen_cb.addWidget(self.cb_filtre_cd, 0, 2)
        duzen_cb.addWidget(self.cb_filtre_pt100, 1, 0)
        duzen_cb.addWidget(self.cb_filtre_isitici, 1, 1)
        duzen_cb.addWidget(self.cb_filtre_imak, 1, 2)

        grup_checkboxlar.setLayout(duzen_cb)
        self.layout.addWidget(grup_checkboxlar)

        self.layout.addWidget(grup_tarih)

        self.btn_ara = QPushButton("Filtrele ve Tabloya Getir")
        self.btn_ara.setStyleSheet("background-color: #2980B9; color: white; font-weight: bold; padding: 10px;")
        self.btn_ara.clicked.connect(self.arama_yap)
        self.layout.addWidget(self.btn_ara)

    def arama_yap(self):
        try:
            self.parent.tablo.setSortingEnabled(False)

            if self.df.empty:
                self.df = pd.read_excel(self.excel_yolu, dtype=str, header=0).fillna("")
                self.temiz_kolonlar = {turkce_karakter_temizleme(str(col)).strip().upper(): col for col in
                                       self.df.columns}

            df_filtrelenmis = self.df.copy()

            def dna_cikar(metin):
                t = turkce_karakter_temizleme(str(metin)).upper()
                for isaret in ["-", ":", "/", "\\", "_", ",", ".", "MM"]:
                    t = t.replace(isaret, " ")
                return set([k.strip() for k in t.split() if k.strip()])

            # 1. Metin / ComboBox Filtreleri
            for temiz_baslik, kutu in self.arama_kutu_sozlugu.items():
                aranan_metin = turkce_karakter_temizleme(kutu.currentText().strip()).upper()
                if aranan_metin and temiz_baslik in self.temiz_kolonlar:
                    gercek_kolon = self.temiz_kolonlar[temiz_baslik]
                    aranan_kelimeler = aranan_metin.split()

            def iceriyor_mu(hucre_verisi):
                hucre_str = turkce_karakter_temizleme(str(hucre_verisi)).upper()
                return all(kel in hucre_str for kel in aranan_kelimeler)

            mask = df_filtrelenmis[gercek_kolon].apply(iceriyor_mu)
            df_filtrelenmis = df_filtrelenmis[mask]

            # 2. Sıra No (Kod) Filtresi
            sira_girdisi = self.sira_no_kutusu.text().strip()
            if sira_girdisi.isdigit():
                aranan_sira_kod = sira_girdisi.zfill(4)
                sira_kolonu = self.temiz_kolonlar.get("ACIKLAMA KOD", self.temiz_kolonlar.get("ACIKLAMA SIRA KOD"))
                if sira_kolonu:
                    temiz_sutun = df_filtrelenmis[sira_kolonu].astype(str).str.replace(".0", "",
                                                                                       regex=False).str.strip().str.zfill(
                        4)
                    df_filtrelenmis = df_filtrelenmis[temiz_sutun == aranan_sira_kod]

            # --- YENİ: CHECKBOX İLE FİLTRELEME ---
            # Satırdaki tüm hücreleri dev bir metne çevirip içinde arıyoruz
            satirlar_metni = df_filtrelenmis.astype(str).apply(lambda x: ' '.join(x).upper(), axis=1)
            satirlar_metni = satirlar_metni.apply(turkce_karakter_temizleme)

            if self.cb_filtre_guduk.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("GUDUK", regex=False)]
            if self.cb_filtre_marine.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("MARINE", regex=False)]
            if self.cb_filtre_cd.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("CEKTIRME DELIKSIZ", regex=False)]
            if self.cb_filtre_pt100.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("PT100", regex=False)]
            if self.cb_filtre_isitici.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("ISITICI", regex=False)]
            if self.cb_filtre_imak.isChecked():
                df_filtrelenmis = df_filtrelenmis[satirlar_metni.str.contains("IMAK", regex=False)]

            # 3. Tarih Filtresi
            if getattr(self, "cb_tarih_kullan", None) and self.cb_tarih_kullan.isChecked():
                tarih_kolonu = self.temiz_kolonlar.get("KAYIT TARIHI", self.temiz_kolonlar.get("KAYIT TARİHİ"))
                if tarih_kolonu:
                    b_tar = pd.to_datetime(self.bas_tarih.date().toString("yyyy-MM-dd"))
                    b_bit = pd.to_datetime(self.bit_tarih.date().toString("yyyy-MM-dd"))
                    parsed_dates = pd.to_datetime(df_filtrelenmis[tarih_kolonu], dayfirst=True, errors='coerce')
                    df_filtrelenmis = df_filtrelenmis[(parsed_dates >= b_tar) & (
                                parsed_dates <= b_bit + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

            # 4. Tabloya Basım (Dinamik)
            self.parent.tablo.setRowCount(0)
            id_kolonu = self.temiz_kolonlar.get("KOD", None)

            for index, row in df_filtrelenmis.iterrows():
                mevcut_satir = self.parent.tablo.rowCount()
                self.parent.tablo.insertRow(mevcut_satir)

                if id_kolonu and pd.notna(row[id_kolonu]):
                    self.parent.tablo.setItem(mevcut_satir, 0, QTableWidgetItem(str(row[id_kolonu]).strip()))

                for c in range(self.parent.tablo.columnCount()):
                    tablo_baslik_adi = turkce_karakter_temizleme(
                        self.parent.tablo.horizontalHeaderItem(c).text()).upper()
                    if tablo_baslik_adi in self.temiz_kolonlar:
                        gercek_kolon = self.temiz_kolonlar[tablo_baslik_adi]
                        veri = row[gercek_kolon]
                        if pd.notna(veri):
                            self.parent.tablo.setItem(mevcut_satir, c, QTableWidgetItem(str(veri).strip()))

            self.parent.tablo.setSortingEnabled(True)
            self.parent.satir_sayaci.setText(f"Bulunan Kayıt: {self.parent.tablo.rowCount()}")

            if self.parent.tablo.rowCount() == 0:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.information(self, "Bulunamadı", "Aradığınız kriterlere uygun kayıt bulunamadı.")

            self.accept()

        except Exception as e:
            self.parent.tablo.setSortingEnabled(True)
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Hata", f"Filtreleme sırasında hata oluştu:\n{e}")