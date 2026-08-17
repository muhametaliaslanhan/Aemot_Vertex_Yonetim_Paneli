from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QSpacerItem, QSizePolicy,
    QScrollArea, QMessageBox, QCompleter, QDialog, QDateEdit, QFormLayout, QMenu
)
from PyQt6.QtCore import Qt, QTimer, QDateTime,QDate
from PyQt6.QtGui import QPixmap
import os
import stat
import shutil
import tempfile
import pandas as pd
import xlwings as xw

from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir,kaynak_yolu
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.motor_verileri import turkce_karakter_temizleme
from arayuz.mantik import oturum

class RotorluMilArayuzu(QWidget):
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

        # --- YENİ: ÜST BUTONLAR (EN TEPEDE OLACAK) ---
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

        baslik = QLabel("ROTORLU MİL OLUŞTURUCU")
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
             "400", "450", "500", "630"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(["","2", "4", "6", "8", "10", "12", "2/4", "4/2"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(["","S", "M", "L", "S-M", "S-L", "M-L", "S-M-L"])

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

        # --- GRUP 2: Rotorlu Mile Özel Girdiler ---
        grup_ozel = QGroupBox("Rotorlu Mil Detayları")
        duzen_ozel = QGridLayout()

        # START OF CHANGES
        # Create a container widget to hold either the LineEdit or ComboBox
        self.paket_boyu_container = QWidget()
        self.paket_boyu_layout = QVBoxLayout(self.paket_boyu_container)
        self.paket_boyu_layout.setContentsMargins(0, 0, 0, 0)

        self.paket_boyu_input = QLineEdit()
        self.paket_boyu_input.setPlaceholderText("Örn: 950")

        self.paket_boyu_kutusu = QComboBox()
        self.paket_boyu_kutusu.addItems(["", "A", "B", "C", "D", "H", "K", "P", "X", "Y", "Z", "Q", "A-V"])
        self.paket_boyu_kutusu.setEditable(True)
        self.paket_boyu_kutusu.hide()  # Hidden by default

        self.paket_boyu_layout.addWidget(self.paket_boyu_input)
        self.paket_boyu_layout.addWidget(self.paket_boyu_kutusu)

        # Method to handle the switch
        def handle_versiyon_degisimi(versiyon_text):
            if versiyon_text == "ESKİ_GÖVDE":
                self.paket_boyu_input.hide()
                self.paket_boyu_kutusu.show()
                # Ensure the current text is synced or cleared appropriately based on requirements
                self.paket_boyu_kutusu.setCurrentText("")
            else:
                self.paket_boyu_kutusu.hide()
                self.paket_boyu_input.show()
                self.paket_boyu_input.clear()

        self.versiyon_kutusu.currentTextChanged.connect(handle_versiyon_degisimi)

        # Call it once to set initial state based on default selection
        handle_versiyon_degisimi(self.versiyon_kutusu.currentText())
        # END OF CHANGES

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        self.resim_no_input = QLineEdit()

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Örn: BAKIR BARALI")

        duzen_ozel.addWidget(QLabel("Paket Boyu (PB) / Tipi:"), 0, 0)
        duzen_ozel.addWidget(self.paket_boyu_container, 0, 1)  # Use the container here
        duzen_ozel.addWidget(QLabel("Revizyon No:"), 0, 2)
        duzen_ozel.addWidget(self.revizyon_kutusu, 0, 3)

        duzen_ozel.addWidget(QLabel("Resim No:"), 1, 0)
        duzen_ozel.addWidget(self.resim_no_input, 1, 1)
        duzen_ozel.addWidget(QLabel("Ek Açıklama:"), 1, 2)
        duzen_ozel.addWidget(self.ek_aciklama_input, 1, 3)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        # --- GRUP 3: Onay Kutuları (Checkbox) ---
        grup_onaylar = QGroupBox("Ek Donanımlar")
        duzen_onaylar = QGridLayout()

        self.cb_guduk = QCheckBox("GÜDÜK")
        self.cb_derin = QCheckBox("DERİN")
        self.cb_onden_sabit = QCheckBox("ÖNDEN SABİT")
        self.cb_arkadan_sabit = QCheckBox("ARKADAN SABİT")
        self.cb_cektirme_deliksiz = QCheckBox("ÇEKTİRME DELİKSİZ")
        self.cb_senkron = QCheckBox("SENKRON")
        self.cb_miknatisli = QCheckBox("MIKNATISLI")
        self.cb_bakir_barali = QCheckBox("BAKIR BARALI")

        duzen_onaylar.addWidget(self.cb_guduk, 0, 0)
        duzen_onaylar.addWidget(self.cb_derin, 0, 1)
        duzen_onaylar.addWidget(self.cb_onden_sabit, 0, 2)
        duzen_onaylar.addWidget(self.cb_arkadan_sabit, 0, 3)

        duzen_onaylar.addWidget(self.cb_cektirme_deliksiz, 1, 0)
        duzen_onaylar.addWidget(self.cb_senkron, 1, 1)
        duzen_onaylar.addWidget(self.cb_miknatisli, 1, 2)
        duzen_onaylar.addWidget(self.cb_bakir_barali, 1, 3)

        grup_onaylar.setLayout(duzen_onaylar)
        form_ana_duzen.addWidget(grup_onaylar)


        scroll_area.setWidget(form_widget)
        ana_duzen.addWidget(scroll_area)
        self.form_alani_widgeti = scroll_area

        # ---------------- 3. ÜRET BUTONU ----------------
        self.btn_uret = QPushButton("LİSTEYE EKLE (KOD ÜRET)")
        self.btn_uret.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 15px; border-radius: 4px; font-size: 14px;")
        self.btn_uret.clicked.connect(self.uret_calistir)
        ana_duzen.addWidget(self.btn_uret)

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

        # arama_duzeni.addWidget(self.btn_ara) satırının altına şunu da ekle:


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
        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR (ROTORLU MIL)")
        self.aktar_butonu.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.aktar_butonu.clicked.connect(self.excele_aktar)
        ana_duzen.addWidget(self.aktar_butonu)

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
            nesne.clear()
        for nesne in self.findChildren(QCheckBox):
            nesne.setChecked(False)

        # Clear specific inputs
        self.paket_boyu_input.clear()
        self.paket_boyu_kutusu.setCurrentIndex(0)

        # Ekstra: Arama kutularını ve sayacı da sıfırla
        self.id_arama_kutusu.clear()
        self.arama_kutusu.clear()
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
        # Artık ilk boşluğu değil, sistemdeki EN BÜYÜK numarayı bulup +1 yapacağız.
        kullanilan_nolar = [0] # Liste boş kalmasın diye 0 ekliyoruz

        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = "ROTORLU MIL"
            for sayfa in excel_dosyasi.sheet_names:
                if sayfa.strip().upper() == "ROTORLU MIL":
                    hedef_sayfa = sayfa
                    break

            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna("")
            for idx, row in df.iterrows():
                for val in row.values:
                    val_str = str(val).strip().upper()
                    if "RM-" in val_str:
                        parcalar = val_str.split("-")
                        for parca in parcalar:
                            if parca.isdigit() and len(parca) == 4:
                                kullanilan_nolar.append(int(parca))
        except Exception as e:
            pass

        # Tablodaki numaraları da dahil et
        for row_idx in range(self.tablo.rowCount()):
            try:
                tablo_id = self.tablo.item(row_idx, 0).text().strip().upper()
                if "RM-" in tablo_id:
                    parcalar = tablo_id.split("-")
                    for parca in parcalar:
                        if parca.isdigit() and len(parca) == 4:
                            kullanilan_nolar.append(int(parca))
            except:
                pass

        # 1'den başlayıp kullanılmayan ilk boşluğu (örn: 85 ve 87 doluysa 86'yı) bulur.
        for i in range(1, 100000):
            if i not in kullanilan_nolar:
                return i

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip().upper())


        if not aranan_id and not aranan_aciklama:
            return

        guncel_yol = ayar_excel_yolunu_getir()
        self.tablo.setRowCount(0)

        # Hangi sayfada olduğumuzu akıllıca tespit eden mekanizma
        gercek_sayfa_adi = "ROTORLU MIL"
        try:
            excel_dosyasi = pd.ExcelFile(guncel_yol)
            for sayfa in excel_dosyasi.sheet_names:
                s_up = sayfa.upper().strip()
                if "ROTORLU" in s_up and ("MIL" in s_up or "MİL"):
                    gercek_sayfa_adi = sayfa
                    break
        except:
            pass

        try:
            # Artık A1'den başlıyor, başlık arama döngülerini çöpe attık
            df = pd.read_excel(guncel_yol, sheet_name=gercek_sayfa_adi, header=0, dtype=str).fillna("")
            df_veri = df.copy()

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
            dialog = RotorGelismisArama(self, excel_yolu)
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Gelişmiş arama açılamadı:\n{e}")

    def excele_aktar(self):
        gercek_sayfa_adi = "ROTORLU MIL"
        guncel_yol = ayar_excel_yolunu_getir()
        excel_yolu = os.path.abspath(guncel_yol)

        if self.tablo.rowCount() == 0:
            QMessageBox.warning(self, "Uyarı", "Tabloda aktarılacak veri yok!")
            return

        # 1. PANDAS İLE KOPYA KONTROLÜ (Ağdaki dosyayı sadece okuduğumuz için kilitlenmez)
        try:
            df = pd.read_excel(excel_yolu, sheet_name=None, header=None, dtype=str)
            for sayfa in df.keys():
                if "ROTORLU" in sayfa.upper() and ("MIL" in sayfa.upper() or "MİL" in sayfa.upper()):
                    gercek_sayfa_adi = sayfa
                    break

            sayfa_df = df[gercek_sayfa_adi].fillna("")

            baslik_satiri = 7
            for i in range(20):
                if "KOD" in " ".join(list(sayfa_df.iloc[i])).upper():
                    baslik_satiri = i
                    break
            sayfa_df.columns = [str(c).strip().upper() for c in sayfa_df.iloc[baslik_satiri]]

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

            if cakisan_satirlar:
                satir_numaralari = ", ".join([str(r + 1) for r in cakisan_satirlar])
                cevap_kopya = QMessageBox.question(
                    self, "KOPYA ROTOR UYARISI!",
                    f"Tablodaki {len(cakisan_satirlar)} adet kayıt Excel'de zaten mevcut!\n\n"
                    f"• EVET: Kopyaları tablodan kaldır ve sadece yenileri kaydet.\n"
                    f"• İPTAL: İşlemi tamamen durdur.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel
                )
                if cevap_kopya == QMessageBox.StandardButton.Cancel:
                    return
                elif cevap_kopya == QMessageBox.StandardButton.Yes:
                    for r in sorted(cakisan_satirlar, reverse=True):
                        self.tablo.removeRow(r)
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

        # ---------------------------------------------------------
        # 3. YENİ SİSTEM: LOKALDE DEĞİŞTİRİP GLOBALE ATMA (SHADOW COPY)
        # ---------------------------------------------------------
        try:
            # 3.1 Ağdaki dosyayı bilgisayarın geçici belleğine (Temp) kopyala
            temp_klasor = tempfile.gettempdir()
            gecici_excel = os.path.join(temp_klasor, "gecici_aemot_rotorlu_mil.xlsx")

            # Dosyayı lokalimize kopyalıyoruz
            shutil.copy2(excel_yolu, gecici_excel)

            # Lokaldeki dosyanın salt okunur kilidini %100 kırıyoruz (çünkü dosya artık bizim)
            os.chmod(gecici_excel, stat.S_IWRITE)

            # 3.2 xlwings ile lokaldeki dosyayı açıyoruz (Asla donmaz, gizli hata vermez)
            app = xw.App(visible=False)
            app.display_alerts = False

            wb = app.books.open(gecici_excel)
            sheet = wb.sheets[gercek_sayfa_adi]


            # 1. Satır H ve I Sütunlarına başlıkları yaz (Artık A1'den başlıyor)
            sheet.range('H1').value = "KAYIT TARİHİ"
            sheet.range('I1').value = "KAYDI EKLEYEN KİŞİ"
            try:
                sheet.range('A1:I1').api.Font.Bold = True
            except:
                pass

            # A sütunundaki son dolu satırı bul
            son_satir = sheet.range('A1048576').end('up').row
            if son_satir < 1: son_satir = 1

            aktarilacak_veriler = []
            for r in range(self.tablo.rowCount()):
                satir_verisi = []
                for c in range(8):
                    item = self.tablo.item(r, c)
                    satir_verisi.append(item.text() if item else "")

                # SATIRIN SONUNA KULLANICIYI ZIMBALA
                satir_verisi.append(oturum.tam_isim())
                aktarilacak_veriler.append(satir_verisi)

            if aktarilacak_veriler:
                baslama_satiri = son_satir + 1
                bitis_satiri = baslama_satiri + len(aktarilacak_veriler) - 1

                # I sütununa kadar genişlettik
                hedef_alan = sheet.range(f'A{baslama_satiri}:I{bitis_satiri}')

                # İŞTE SİHİR BURADA: Bu alanı Excel'de zorla "Metin" formatına çeviriyoruz ('@' metin demektir)
                hedef_alan.number_format = '@'

                # Şimdi verileri basıyoruz. Artık 0000 veya 063 asla silinmeyecek!
                hedef_alan.value = aktarilacak_veriler

            # Lokal dosyayı kaydet ve kapat
            wb.save(gecici_excel)
            wb.close()
            app.quit()

            # 3.3 Lokaldeki GÜNCEL dosyayı ağdaki ORİJİNAL dosyanın üstüne yaz!
            try:
                shutil.copy2(gecici_excel, excel_yolu)
                QMessageBox.information(self, "Başarılı",
                                        "Kayıtlar ağdaki Excel'e (9. satırdan itibaren) kusursuzca eklendi!")
                self.tablo.setRowCount(0)  # Aktarılanları tablodan temizle

            except PermissionError:
                # EĞER AĞDA BİRİSİ DOSYAYI AÇIK TUTUYORSA
                QMessageBox.critical(self, "DOSYA BAŞKA BİLGİSAYARDA AÇIK!",
                                     f"Veriler hazırlandı ancak AĞDAKİ Excel dosyası şu an başka bir bilgisayarda AÇIK olduğu için kayıt tamamlanamadı.\n\n"
                                     f"Lütfen diğer bilgisayarlardan Excel'in kapatılmasını isteyin ve tablodaki verileriniz silinmeden tekrar 'Excel'e Aktar' butonuna basın.")
            except Exception as e:
                QMessageBox.critical(self, "Ağ Hatası", f"Dosya ağa gönderilirken hata oluştu:\n{e}")

        except Exception as e:
            try:
                app.quit()
            except:
                pass
            QMessageBox.critical(self, "Lokal İşlem Hatası", f"Lokal işleme sırasında bir hata oluştu:\n{e}")

    def dna_bazli_sira_no_bul(self, sayfa_adi, aranan_grup, saf_dna_aciklama2):
        # TARİH KESİNLİKLE DAHİL DEĞİLDİR! Sadece Grup Kodu ve Açıklama2 BİREBİR uyuşuyorsa aynı kodu al.

        # DÜZELTME: Excel'e elle girilen kayıtlarda TR karakterler (Ğ Ü Ç Ö Ş İ) tutarsız olabiliyor.
        # Karşılaştırmadan önce her iki tarafı da normalize ediyoruz ki sırf harf farkından
        # dolayı aynı parçaya farklı sıra no verilmesin.
        aranan_grup = turkce_karakter_temizleme(str(aranan_grup)).strip().upper()
        saf_dna_aciklama2 = turkce_karakter_temizleme(str(saf_dna_aciklama2)).strip().upper()

        # 1. ARAYÜZ TABLOSUNA BAK
        for row_idx in range(self.tablo.rowCount()):
            tablo_grup = turkce_karakter_temizleme(str(self.tablo.item(row_idx, 2).text())).strip().upper()
            tablo_aciklama2 = turkce_karakter_temizleme(str(self.tablo.item(row_idx, 3).text())).strip().upper()

            if tablo_grup == aranan_grup and tablo_aciklama2 == saf_dna_aciklama2:
                try:
                    sira_no = int(self.tablo.item(row_idx, 5).text().strip())  # 5. index AÇIKLAMA S.NO sütunudur
                    return sira_no
                except:
                    pass

        # 2. EXCEL GEÇMİŞİNE BAK
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = sayfa_adi
            for sayfa in excel_dosyasi.sheet_names:
                if "ROTORLU" in sayfa.upper() and "MIL" in sayfa.upper():
                    hedef_sayfa = sayfa
                    break

            # header=0 ile direkt A1'den okuyoruz
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=0, dtype=str).fillna(
                "")
            df.columns = [str(c).strip().upper() for c in df.columns]

            for index, row in df.iterrows():
                excel_grup = turkce_karakter_temizleme(str(row.get('GRUP', ''))).strip().upper()
                excel_aciklama2 = turkce_karakter_temizleme(
                    str(row.get('AÇIKLAMA2', row.get('EK AÇIKLAMA', '')))).strip().upper()

                # DNA EŞLEŞTİ Mİ? (Tarihle alakası yok!)
                if excel_grup == aranan_grup and excel_aciklama2 == saf_dna_aciklama2:

                    # TAM SENİN SÜTUN ADIN: AÇIKLAMA S.NO
                    sira_kolonu = 'AÇIKLAMA S.NO'

                    # Eğer bulamazsa yedeğe baksın
                    if sira_kolonu not in df.columns:
                        for col in df.columns:
                            if "S.NO" in col or "SIRA NO" in col:
                                sira_kolonu = col
                                break
                    try:
                        sira_no = int(float(str(row.get(sira_kolonu, '')).replace("'", "").strip()))
                        return sira_no
                    except:
                        pass
        except Exception as e:
            pass

        return None

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

        # Get paket_boyu from the appropriate input based on version
        if versiyon == "ESKİ_GÖVDE":
            paket_boyu = self.paket_boyu_kutusu.currentText().strip().upper()
        else:
            paket_boyu = self.paket_boyu_input.text().strip().upper()

        revizyon = self.revizyon_kutusu.currentText()

        # =========================================================================
        # --- ESKİ GÖVDE: GÖVDE BOYU VE PAKET BOYU BAĞLILIK KONTROLÜ ---
        # =========================================================================
        if versiyon == "ESKİ_GÖVDE":
            # XOR check: One is empty, the other is not
            if bool(govde_boyu) != bool(paket_boyu):
                # Özel Uyarı Penceresi
                mesaj_kutusu = QMessageBox(self)
                mesaj_kutusu.setIcon(QMessageBox.Icon.Warning)
                mesaj_kutusu.setWindowTitle("Kural İhlali: Bağlı Seçimler")
                mesaj_kutusu.setText(
                    "ESKİ GÖVDE versiyonunda 'Gövde Boyu' ve 'Paket Boyu' BİRLİKTE kullanılmalıdır.\n\n"
                    "Ya ikisini de BOŞ bırakın (Örn: 071 4 ROTORLU MIL) \n"
                    "Ya da ikisini de DOLDURUN (Örn: 071 M4B ROTORLU MIL).\n\n"
                    "Ne yapmak istersiniz?"
                )

                # Custom buttons
                btn_temizle = mesaj_kutusu.addButton("İkisini de Temizle", QMessageBox.ButtonRole.ActionRole)
                btn_doldur = mesaj_kutusu.addButton("Doldurmaya Geri Dön", QMessageBox.ButtonRole.RejectRole)

                mesaj_kutusu.exec()

                if mesaj_kutusu.clickedButton() == btn_temizle:
                    self.govde_boyu_kutusu.setCurrentIndex(0)
                    self.paket_boyu_kutusu.setCurrentIndex(0)

                return  # Stop execution until the logic is correct

        # YENİ: Resim No Boş İse "0000" Yap
        resim_no_girdi = self.resim_no_input.text().strip()
        ham_resim_no_bos_mu = False
        if not resim_no_girdi:
            resim_no_girdi = "0000"
            ham_resim_no_bos_mu = True

        resim_no_tablo_icin = resim_no_girdi  # Artık tek tırnağa gerek yok, saf 0000

        # --- BOŞ ALAN UYARISI ---
        bos_alanlar = []
        if not grup: bos_alanlar.append("Grup")
        if not kutup: bos_alanlar.append("Kutup")

        # For Premium, enforce govde_boyu and paket_boyu
        if versiyon == "PREMIUM":
            if not govde_boyu: bos_alanlar.append("Gövde Boyu")
            if not paket_boyu: bos_alanlar.append("Paket Boyu")

        if bos_alanlar:
            cevap = QMessageBox.question(self, "Eksik Bilgi",
                                         f"Şu alanlar boş:\n{', '.join(bos_alanlar)}\n\nYine de üretip tabloya eklemek istiyor musunuz?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        ham_ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
        ek_aciklama = ham_ek_aciklama.replace("ROTORLU MIL", "").replace("ROTORLU MİL", "").strip()

        guduk = self.cb_guduk.isChecked()
        derin = self.cb_derin.isChecked()
        onden = self.cb_onden_sabit.isChecked()
        arkadan = self.cb_arkadan_sabit.isChecked()
        cektirme = self.cb_cektirme_deliksiz.isChecked()
        senkron = self.cb_senkron.isChecked()
        miknatisli = self.cb_miknatisli.isChecked()
        bakir_barali = self.cb_bakir_barali.isChecked()

        # ---------------- 2. KUSURSUZ DNA (AÇIKLAMA 2) OLUŞTURMA ----------------
        dna_parcalari = []
        ozel_durum_var_mi_girdi = guduk or derin or onden or arkadan or cektirme or senkron or miknatisli or bakir_barali or ham_resim_no_bos_mu

        if versiyon == "ESKİ_GÖVDE":
            # ESKİ GÖVDE: Gövde Boyu, Kutup, Paket Boyu Kombinasyonu
            if govde_boyu and paket_boyu:
                # İkisi de doluysa: M4A, S2B gibi birleştir
                ebat_str = f"{govde_boyu}{kutup}{paket_boyu}".strip()
                dna_parcalari.append(ebat_str)
            else:
                # İkisi de boşsa sadece kutup (Örn: "4")
                if kutup:
                    dna_parcalari.append(kutup)

            if guduk: dna_parcalari.append("GUDUK")
            if derin: dna_parcalari.append("DERIN")
            if onden: dna_parcalari.append("ONDEN SABIT")
            if arkadan: dna_parcalari.append("ARKADAN SABIT")
            if cektirme: dna_parcalari.append("CEKTIRME DELIKSIZ")
            if senkron: dna_parcalari.append("SENKRON")
            if miknatisli: dna_parcalari.append("MIKNATISLI")
            if bakir_barali: dna_parcalari.append("BAKIR BARALI")

            if ek_aciklama:
                dna_parcalari.append(ek_aciklama)

            dna_parcalari.append("ROTORLU MIL")

            # Açıklama 2 (Resim No KESİNLİKLE YOK)
            saf_dna_aciklama2 = " ".join(dna_parcalari)
            saf_dna_aciklama2 = " ".join(saf_dna_aciklama2.split())

            # Ana Açıklama (Resim no doluysa ROTORLU MIL'in önüne eklenir)
            uretilen_aciklama = f"{grup} {saf_dna_aciklama2}".strip()
            if not ozel_durum_var_mi_girdi and resim_no_girdi not in ["0000", "00000"]:
                uretilen_aciklama = uretilen_aciklama.replace("ROTORLU MIL", f"{resim_no_girdi} ROTORLU MIL").strip()

        else:
            # PREMIUM FORMATI
            if ozel_durum_var_mi_girdi:
                dna_parcalari.append(f"-{kutup}/{govde_boyu}")
            else:
                dna_parcalari.append(f"-{kutup}")

            if guduk: dna_parcalari.append("GUDUK")
            if derin: dna_parcalari.append("DERIN")
            if onden: dna_parcalari.append("ONDEN SABIT")
            if arkadan: dna_parcalari.append("ARKADAN SABIT")
            if cektirme: dna_parcalari.append("CEKTIRME DELIKSIZ")
            if senkron: dna_parcalari.append("SENKRON")
            if miknatisli: dna_parcalari.append("MIKNATISLI")
            if bakir_barali: dna_parcalari.append("BAKIR BARALI")

            # Premium'da kullanıcının girdiği sayısal veya metinsel veriyi PB:950 MM olarak yazar
            if paket_boyu:
                dna_parcalari.append(f"ROTORLU MIL PB:{paket_boyu} MM")
            else:
                dna_parcalari.append(f"ROTORLU MIL")

            # Açıklama 2 (Resim No ve Revizyon YOK)
            saf_dna_aciklama2 = " ".join(dna_parcalari)
            saf_dna_aciklama2 = " ".join(saf_dna_aciklama2.split())
            if ek_aciklama:
                saf_dna_aciklama2 = f"{saf_dna_aciklama2} {ek_aciklama}".strip()

            # Ana Açıklama (Revizyon ve Resim No Sona Eklenir)
            uretilen_aciklama = f"{grup} {saf_dna_aciklama2} - R:{revizyon}".strip()
            if not ozel_durum_var_mi_girdi and resim_no_girdi not in ["0000", "00000"]:
                uretilen_aciklama += f" - {resim_no_girdi}"

        # Ekstra boşlukları temizle
        uretilen_aciklama = " ".join(uretilen_aciklama.split())

        # ---------------- 3. DNA TESTİ İLE SIRA NO BULMA ----------------
        dna_sira = self.dna_bazli_sira_no_bul(sayfa_adi="ROTORLU MIL", aranan_grup=grup,
                                              saf_dna_aciklama2=saf_dna_aciklama2)

        from arayuz.mantik.sira_no_dialog import sira_no_secim_yap
        secilen_sira = sira_no_secim_yap(self, dna_sira, uretilen_aciklama, sayfa_adi="ROTORLU MIL", id_kodu="RM")
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
        id_resim_no = "00000" if ozel_durum_var_mi_girdi else resim_no_girdi.zfill(5)

        uretilen_id = f"YMM{grup}-RM-{aciklama_kod_str}-{id_revizyon}-{id_resim_no}"


        if versiyon == "ESKİ_GÖVDE":
            # Eski gövdede revizyon eklenmez (- R:01 YOK) ve resim no 0000'dan farklıysa ROTORLU MIL öncesine gelir
            if not ozel_durum_var_mi_girdi and resim_no_girdi != "0000":
                uretilen_aciklama = f"{grup} {saf_dna_aciklama2}".replace("ROTORLU MIL",
                                                                          f"{resim_no_girdi} ROTORLU MIL").strip()
            else:
                uretilen_aciklama = f"{grup} {saf_dna_aciklama2}".strip()
            uretilen_aciklama = " ".join(uretilen_aciklama.split())
        else:
            uretilen_aciklama = f"{grup} {saf_dna_aciklama2} - R:{revizyon}".strip()
            uretilen_aciklama = " ".join(uretilen_aciklama.split())

        # --- KOPYA KALKANI ---
        for r in range(self.tablo.rowCount()):
            tab_id = self.tablo.item(r, 0).text().strip()
            tab_aciklama = self.tablo.item(r, 1).text().strip()
            if tab_id == uretilen_id and tab_aciklama == uretilen_aciklama:
                QMessageBox.warning(self, "KOPYA ENGELİ", "Bu kayıt ZATEN LİSTEDE var!")
                return

        # ---------------- 5. TABLOYA YAZDIRMA ----------------
        bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

        # resim_no_girdi kısmı boş bırakıldıysa artık içinde "0000" var, tabloya o şekilde basacak.
        yeni_kayit = [uretilen_id, uretilen_aciklama, grup, saf_dna_aciklama2, resim_no_tablo_icin, sira_no_str,
                      aciklama_kod_str, bugunun_tarihi]

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


class RotorGelismisArama(QDialog):
    def __init__(self, parent, excel_yolu):
        super().__init__(parent)
        self.setWindowTitle("Rotorlu Mil Gelişmiş Arama")
        self.setMinimumWidth(650)
        self.parent = parent
        self.excel_yolu = excel_yolu

        self.layout = QVBoxLayout(self)

        # --- 1. EXCEL VERİSİNİ OKU ---
        gercek_sayfa = "ROTORLU MIL"
        try:
            dosya = pd.ExcelFile(self.excel_yolu)
            for s in dosya.sheet_names:
                if "ROTORLU" in s.upper() and ("MIL" in s.upper() or "MİL" in s.upper()):
                    gercek_sayfa = s
                    break
            self.df = pd.read_excel(self.excel_yolu, sheet_name=gercek_sayfa, header=0, dtype=str).fillna("")
            self.df.columns = [str(c).strip().upper() for c in self.df.columns]
            self.temiz_kolonlar = {str(col): col for col in self.df.columns}
        except:
            self.df = pd.DataFrame()
            self.temiz_kolonlar = {}

        # --- 2. FORM ALANI (BİREBİR ANA FORMUN KOPYASI) ---
        grup_form = QGroupBox("Rotorlu Mil Parametreleriyle Filtrele")
        form_layout = QGridLayout()

        # Kutular
        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["", "ESKİ_GÖVDE", "PREMIUM"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["", "063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315",
             "355", "400", "450", "500", "630"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(["", "2", "4", "6", "8", "10", "12", "2/4", "4/2"])

        self.govde_boyu_kutusu = QComboBox()
        self.govde_boyu_kutusu.addItems(["", "S", "M", "L", "S-M", "S-L", "M-L", "S-M-L"])

        # --- START CHANGED FOR GELISMIS ARAMA ---
        self.paket_boyu_container = QWidget()
        self.paket_boyu_layout = QVBoxLayout(self.paket_boyu_container)
        self.paket_boyu_layout.setContentsMargins(0, 0, 0, 0)

        self.paket_boyu_input = QLineEdit()
        self.paket_boyu_input.setPlaceholderText("Örn: 950")

        self.paket_boyu_kutusu = QComboBox()
        self.paket_boyu_kutusu.addItems(["", "A", "B", "C", "D", "H", "K", "P", "X", "Y", "Z", "Q", "A-V"])
        self.paket_boyu_kutusu.setEditable(True)
        self.paket_boyu_kutusu.hide()

        self.paket_boyu_layout.addWidget(self.paket_boyu_input)
        self.paket_boyu_layout.addWidget(self.paket_boyu_kutusu)

        def handle_arama_versiyon_degisimi(versiyon_text):
            if versiyon_text == "ESKİ_GÖVDE":
                self.paket_boyu_input.hide()
                self.paket_boyu_kutusu.show()
                self.paket_boyu_kutusu.setCurrentText("")
            else:
                self.paket_boyu_kutusu.hide()
                self.paket_boyu_input.show()
                self.paket_boyu_input.clear()

        self.versiyon_kutusu.currentTextChanged.connect(handle_arama_versiyon_degisimi)
        handle_arama_versiyon_degisimi(self.versiyon_kutusu.currentText())
        # --- END CHANGED FOR GELISMIS ARAMA ---

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["", "01", "02", "03", "04", "05"])

        self.resim_no_input = QLineEdit()
        self.resim_no_input.setPlaceholderText("Resim No Ara...")

        self.ek_aciklama_input = QLineEdit()
        self.ek_aciklama_input.setPlaceholderText("Ek Açıklama (Örn: GÜDÜK, BAKIR)")

        self.sira_no_kutusu = QLineEdit()
        self.sira_no_kutusu.setPlaceholderText("Sıra No (Kod) Örn: 1")

        # Yerleştirme
        form_layout.addWidget(QLabel("Versiyon:"), 0, 0)
        form_layout.addWidget(self.versiyon_kutusu, 0, 1)
        form_layout.addWidget(QLabel("Grup (M):"), 0, 2)
        form_layout.addWidget(self.grup_kutusu, 0, 3)

        form_layout.addWidget(QLabel("Kutup:"), 1, 0)
        form_layout.addWidget(self.kutup_kutusu, 1, 1)
        form_layout.addWidget(QLabel("Gövde Boyu:"), 1, 2)
        form_layout.addWidget(self.govde_boyu_kutusu, 1, 3)

        form_layout.addWidget(QLabel("Paket Boyu:"), 2, 0)
        form_layout.addWidget(self.paket_boyu_container, 2, 1)
        form_layout.addWidget(QLabel("Revizyon No:"), 2, 2)
        form_layout.addWidget(self.revizyon_kutusu, 2, 3)

        form_layout.addWidget(QLabel("Resim No:"), 3, 0)
        form_layout.addWidget(self.resim_no_input, 3, 1)
        form_layout.addWidget(QLabel("Ek Açıklama:"), 3, 2)
        form_layout.addWidget(self.ek_aciklama_input, 3, 3)

        form_layout.addWidget(QLabel("Sıra No (Kod):"), 4, 0)
        form_layout.addWidget(self.sira_no_kutusu, 4, 1)

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

        # Combobox'ları yazılabilir yapma
        for kutu in self.findChildren(QComboBox):
            kutu.setEditable(True)
            kutu.lineEdit().mousePressEvent = lambda event, le=kutu.lineEdit(): le.selectAll()

    def arama_yap(self):
        try:
            self.parent.tablo.setSortingEnabled(False)
            df_filtrelenmis = self.df.copy()

            # --- 1. TÜM METİN (DNA) HAVUZU (Filtreleme için) ---
            # DataFrame'in ilk 8 sütununu birleştirerek devasa bir arama havuzu oluşturuyoruz
            tum_metin = pd.Series([""] * len(df_filtrelenmis), index=df_filtrelenmis.index)
            for col_idx in range(min(8, len(df_filtrelenmis.columns))):
                tum_metin += df_filtrelenmis.iloc[:, col_idx].astype(str).str.upper() + " "

            # --- 2. GİRDİLERİ ÇEK ---
            versiyon = self.versiyon_kutusu.currentText().strip().upper()
            grup = self.grup_kutusu.currentText().strip().upper()
            kutup = self.kutup_kutusu.currentText().strip().upper()
            govde_boyu = self.govde_boyu_kutusu.currentText().strip().upper()

            # Paket boyunu versiyona göre doğru kutudan çekiyoruz
            if versiyon == "ESKİ_GÖVDE":
                paket_boyu = self.paket_boyu_kutusu.currentText().strip().upper()
            else:
                paket_boyu = self.paket_boyu_input.text().strip().upper()

            revizyon = self.revizyon_kutusu.currentText().strip().upper()
            resim_no = self.resim_no_input.text().strip().upper()
            ek_aciklama = turkce_karakter_temizleme(self.ek_aciklama_input.text()).strip().upper()
            sira_no = self.sira_no_kutusu.text().strip()

            # --- 3. AKILLI FİLTRELEME ---
            if grup:
                # Grup genelde 2. sütundadır (indeks 2)
                df_filtrelenmis = df_filtrelenmis[df_filtrelenmis.iloc[:, 2].astype(str).str.upper() == grup]

            if resim_no:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(resim_no, regex=False)]

            if kutup:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"-{kutup}", regex=False) | tum_metin.str.contains(f" {kutup} ",
                                                                                              regex=False)]

            if govde_boyu:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"/{govde_boyu}", regex=False) | tum_metin.str.contains(f" {govde_boyu} ",
                                                                                                   regex=False)]

            if paket_boyu:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"PB:{paket_boyu}", regex=False) | tum_metin.str.contains(
                        f"PB: {paket_boyu}", regex=False)]

            if revizyon:
                df_filtrelenmis = df_filtrelenmis[
                    tum_metin.str.contains(f"R:{revizyon}", regex=False) | tum_metin.str.contains(f"R: {revizyon}",
                                                                                                  regex=False) | tum_metin.str.contains(
                        f"-{revizyon}-", regex=False)]

            if ek_aciklama:
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(ek_aciklama, regex=False)]

            if versiyon == "ESKİ_GÖVDE":
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains("-00-", regex=False)]

            if sira_no.isdigit():
                aranan_kod = sira_no.zfill(4)
                # Sıra kod / sıra no sütunlarını tarar
                df_filtrelenmis = df_filtrelenmis[tum_metin.str.contains(aranan_kod, regex=False)]

            if self.cb_tarih_kullan.isChecked():
                b_tar = pd.to_datetime(self.bas_tarih.date().toString("yyyy-MM-dd"))
                b_bit = pd.to_datetime(self.bit_tarih.date().toString("yyyy-MM-dd"))
                # Son sütunlarda tarih arandığı varsayılır
                parsed_dates = pd.to_datetime(df_filtrelenmis.iloc[:, -1], dayfirst=True, errors='coerce')
                df_filtrelenmis = df_filtrelenmis[
                    (parsed_dates >= b_tar) & (parsed_dates <= b_bit + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

            # --- 4. TABLOYA KUSURSUZ VE EKSİKSİZ AKTARIM ---
            # Rotorlu Mil Excel Sıralaması: 0:Kod, 1:Açıklama, 2:Grup, 3:Ek Açıklama, 4:Resim No, 5:Sıra No, 6:Sıra Kod, 7:Tarih
            self.parent.tablo.setRowCount(0)

            for index, row in df_filtrelenmis.iterrows():
                # İlk sütun (Kod) boşsa bu satırı atla
                ilk_hucre = str(row.iloc[0]).strip()
                if not ilk_hucre or ilk_hucre == 'nan' or "KOD" in ilk_hucre.upper():
                    continue

                mevcut_satir = self.parent.tablo.rowCount()
                self.parent.tablo.insertRow(mevcut_satir)

                # Excel'deki satırın ilk 8 sütununu arayüz tablosunun 0-7 sütunlarına birebir basıyoruz
                for col_idx in range(min(8, len(row))):
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