from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QSpacerItem, QSizePolicy,
    QScrollArea, QMessageBox
)
from PyQt6.QtCore import Qt, QTimer, QDateTime
from PyQt6.QtGui import QPixmap
import pandas as pd
import traceback

import os

from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir
from arayuz.mantik.excel_islemleri import toplu_excele_aktar
from arayuz.mantik.motor_verileri import turkce_karakter_temizleme


class RotorPaketArayuzu(QWidget):
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

        # ---------------- 1. ÜST BİLGİ ALANI (HEADER) ----------------
        header_duzen = QHBoxLayout()

        self.logo_etiketi = QLabel()
        self.logo_etiketi.setPixmap(QPixmap("arayuz/varliklar/aemot_logo_gorseli.png"))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)

        baslik = QLabel("ROTOR PAKET OLUŞTURUCU")
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
        grup_temel = QGroupBox("Temel Özellikler")
        duzen_temel = QGridLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["ESKİ_GÖVDE", "PREMIUM"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630"])

        self.kutup_kutusu = QComboBox()
        self.kutup_kutusu.addItems(["2", "4", "6", "8", "10"])

        duzen_temel.addWidget(QLabel("Versiyon:"), 0, 0)
        duzen_temel.addWidget(self.versiyon_kutusu, 0, 1)
        duzen_temel.addWidget(QLabel("Grup (M):"), 0, 2)
        duzen_temel.addWidget(self.grup_kutusu, 0, 3)

        duzen_temel.addWidget(QLabel("Kutup Sayısı:"), 1, 0)
        duzen_temel.addWidget(self.kutup_kutusu, 1, 1)

        grup_temel.setLayout(duzen_temel)
        form_ana_duzen.addWidget(grup_temel)

        # --- GRUP 2: Rotor Pakete Özel Girdiler ---
        grup_ozel = QGroupBox("Rotor Paket Detayları")
        duzen_ozel = QGridLayout()

        self.paket_boyu_input = QLineEdit()
        self.paket_boyu_input.setPlaceholderText("Örn: 47")

        # REVIZYON KUTUSUNU DÜZENLENEBİLİR (YAZILABİLİR) YAPTIK
        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.setEditable(True)
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        self.ek_aciklama_input = QLineEdit()

        duzen_ozel.addWidget(QLabel("Paket Boyu (PB) [mm]:"), 0, 0)
        duzen_ozel.addWidget(self.paket_boyu_input, 0, 1)
        duzen_ozel.addWidget(QLabel("Revizyon No:"), 0, 2)
        duzen_ozel.addWidget(self.revizyon_kutusu, 0, 3)

        duzen_ozel.addWidget(QLabel("Ek Açıklama:"), 1, 0)
        duzen_ozel.addWidget(self.ek_aciklama_input, 1, 1)

        grup_ozel.setLayout(duzen_ozel)
        form_ana_duzen.addWidget(grup_ozel)

        scroll_area.setWidget(form_widget)
        ana_duzen.addWidget(scroll_area)
        self.form_alani_widgeti = scroll_area

        # ---------------- 3. BUTONLAR VE İŞLEMLER ----------------
        buton_duzeni = QHBoxLayout()

        btn_geri = QPushButton("Ana Menü")
        btn_geri.setStyleSheet(
            "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        btn_geri.clicked.connect(lambda: self.ana_pencere.sayfa_degistir(0))

        self.btn_formu_temizle = QPushButton("Formu Sıfırla")
        self.btn_formu_temizle.setStyleSheet(
            "background-color: #95A5A6; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_formu_temizle.clicked.connect(self.formu_temizle)

        self.btn_uret = QPushButton("LİSTEYE EKLE (KOD ÜRET)")
        self.btn_uret.setStyleSheet(
            "background-color: #27AE60; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        self.btn_uret.clicked.connect(self.uret_calistir)

        buton_duzeni.addWidget(btn_geri)
        buton_duzeni.addWidget(self.btn_formu_temizle)
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
        arama_duzeni.addWidget(self.btn_satir_sil)
        arama_duzeni.addWidget(self.btn_tam_ekran)
        arama_duzeni.addWidget(self.satir_sayaci)

        ana_duzen.addLayout(arama_duzeni)

        # ---------------- TABLO ----------------
        self.tablo = QTableWidget(0, 7)
        self.tablo_basliklari = ["#KOD", "AÇIKLAMA", "GRUP", 'AÇIKLAMA2', 'AÇIKLAMA SIRA NO',
                                 'AÇIKLAMA SIRA KOD', "KAYIT TARİHİ"]
        self.tablo.setHorizontalHeaderLabels(self.tablo_basliklari)
        self.tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tablo.setAlternatingRowColors(True)
        ana_duzen.addWidget(self.tablo, stretch=1)

        # ---------------- EXCEL'E AKTAR BUTONU ----------------
        self.aktar_butonu = QPushButton("TÜMÜNÜ EXCEL'E AKTAR (ROTOR PAKET)")
        self.aktar_butonu.setStyleSheet(
            "background-color: #2980B9; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.aktar_butonu.clicked.connect(self.excele_aktar)
        ana_duzen.addWidget(self.aktar_butonu)

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
            if nesne not in [self.arama_kutusu, self.id_arama_kutusu]:
                nesne.clear()
        for nesne in self.findChildren(QCheckBox):
            nesne.setChecked(False)

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
        max_no = 0
        # 1. EXCEL'DEKİ GERÇEK MAX NUMARAYI BUL
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = "ROTOR PAKET"
            for sayfa in excel_dosyasi.sheet_names:
                if "ROTOR" in sayfa.upper() and "PAKET" in sayfa.upper():
                    hedef_sayfa = sayfa
                    break

            # HEADER=NONE diyerek tabloyu kör okuyoruz. Başlık nerede olursa olsun umrumuzda değil.
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=None, dtype=str).fillna("")
            for idx, row in df.iterrows():
                for val in row.values:
                    val_str = str(val).strip().upper()
                    # İçinde RP- olan kodu yakala (Örn: YMM132-RP-0021-01)
                    if "RP-" in val_str:
                        parcalar = val_str.split("-")
                        for parca in parcalar:
                            if len(parca) == 4 and parca.isdigit():
                                no = int(parca)
                                if no > max_no: max_no = no
        except Exception as e:
            print(f"Excel max no bulma hatası: {e}")

        # 2. ARAYÜZ TABLOSUNDAKİ (Henüz kaydedilmemiş) MAX NUMARAYI BUL
        for row_idx in range(self.tablo.rowCount()):
            try:
                # Tablodaki 0. Sütun -> '#KOD' (Örn: YMM132-RP-0021-01)
                tablo_id = self.tablo.item(row_idx, 0).text().strip().upper()
                if "RP-" in tablo_id:
                    parcalar = tablo_id.split("-")
                    for parca in parcalar:
                        if len(parca) == 4 and parca.isdigit():
                            no = int(parca)
                            if no > max_no: max_no = no
            except:
                pass

        return max_no

    def tabloda_arama_yap(self):
        aranan_id = turkce_karakter_temizleme(self.id_arama_kutusu.text().strip()).upper()
        aranan_aciklama = turkce_karakter_temizleme(self.arama_kutusu.text().strip()).upper()

        if len(aranan_id) < 2 and len(aranan_aciklama) < 3:
            return

        guncel_yol = ayar_excel_yolunu_getir()
        self.tablo.setRowCount(0)

        try:
            df = pd.read_excel(guncel_yol, sheet_name="ROTOR PAKET", header=None, dtype=str).fillna("")
            bulunan_kayit_sayisi = 0

            for index, row in df.iterrows():
                tum_satir = " ".join([str(val).strip() for val in row.values])
                tam_havuz = turkce_karakter_temizleme(tum_satir).upper()

                if len(tam_havuz) < 5 or ("KOD" in tam_havuz and "AÇIKLAMA" in tam_havuz):
                    continue

                if "RP" not in tam_havuz:
                    continue

                id_eslesti = True
                aciklama_eslesti = True

                if aranan_id and aranan_id not in tam_havuz:
                    id_eslesti = False

                if aranan_aciklama:
                    aranan_kelimeler = aranan_aciklama.split()
                    if not all(kelime in tam_havuz for kelime in aranan_kelimeler):
                        aciklama_eslesti = False

                if id_eslesti and aciklama_eslesti:
                    mevcut_satir = self.tablo.rowCount()
                    self.tablo.insertRow(mevcut_satir)

                    gercek_veriler = []
                    for val in row.values:
                        temiz_val = str(val).strip()
                        if not temiz_val and len(gercek_veriler) == 0:
                            continue
                        gercek_veriler.append(temiz_val)

                    for i in range(min(len(gercek_veriler), 7)):
                        self.tablo.setItem(mevcut_satir, i, QTableWidgetItem(gercek_veriler[i]))

                    bulunan_kayit_sayisi += 1

            self.satir_sayaci.setText(f"Bulunan Kayıt: {bulunan_kayit_sayisi}")

        except ValueError:
            QMessageBox.warning(self, "Hata", "Excel dosyasında 'ROTOR PAKET' adında bir sayfa bulunamadı!")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Arama sırasında hata:\n{e}")

    def excele_aktar(self):
        cevap = QMessageBox.question(self, "ONAY", "ROTOR PAKET sayfasına aktarılsın mı?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if cevap == QMessageBox.StandardButton.Yes:
            basarili, mesaj = toplu_excele_aktar(self.tablo, sayfa_adi="ROTOR PAKET", baslik_satiri=12)
            if basarili:
                QMessageBox.information(self, "Başarılı", mesaj)
                self.tablo.setRowCount(0)
            else:
                QMessageBox.warning(self, "Hata", mesaj)

    def dna_bazli_sira_no_bul(self, sayfa_adi, aranan_grup, taslak_metin, resim_no_cek=False):

        # =====================================================================
        # SENİN FİKRİN: Metni kelime kelime diziye (array) atıp karşılaştıran motor
        # =====================================================================
        def kelime_dizisi_yap(metin):
            temiz_metin = str(metin).upper()

            # Tüm gereksiz işaretleri tek bir boşluğa çevir
            for isaret in ["-", ":", "/", "\\", "_", ",", ".", "MM"]:
                temiz_metin = temiz_metin.replace(isaret, " ")

            # Fazla boşlukları yut ve DİZİ (Array) yap
            dizi = temiz_metin.split()

            # Sıralamayı umursamaması için alfabetik diz
            return sorted(dizi)

        # =====================================================================

        programin_dizisi = kelime_dizisi_yap(taslak_metin)

        # --- 1. ÖNCE ARAYÜZ TABLOSUNA BAK ---
        for row_idx in range(self.tablo.rowCount()):
            tablo_grup = str(self.tablo.item(row_idx, 2).text()).strip().upper()

            if kelime_dizisi_yap(tablo_grup) == kelime_dizisi_yap(aranan_grup):
                tab_aciklama = str(self.tablo.item(row_idx, 1).text())
                tab_aciklama2 = str(self.tablo.item(row_idx, 3).text())

                tablo_dizisi = kelime_dizisi_yap(f"{tab_aciklama} {tab_aciklama2}")

                if programin_dizisi == tablo_dizisi:
                    sira_no = int(self.tablo.item(row_idx, 4).text().strip())
                    return (sira_no, "") if resim_no_cek else sira_no

        # --- 2. EXCEL GEÇMİŞİNE BAK ---
        try:
            excel_dosyasi = pd.ExcelFile(ayar_excel_yolunu_getir())
            hedef_sayfa = sayfa_adi
            for sayfa in excel_dosyasi.sheet_names:
                if "ROTOR" in sayfa.upper() and "PAKET" in sayfa.upper():
                    hedef_sayfa = sayfa
                    break

            # HEADER KAYMASINI ÇÖZMEK İÇİN TÜM EXCELİ KÖR OKUYORUZ
            df = pd.read_excel(ayar_excel_yolunu_getir(), sheet_name=hedef_sayfa, header=None, dtype=str).fillna("")

            # Excel'in başlık satırını bulma garantisi (İlk 20 satırı tarar)
            baslik_satiri = 7  # Varsayılan 8. satır
            for i in range(20):
                satir_metni = " ".join(list(df.iloc[i])).upper()
                if "GRUP" in satir_metni and "AÇIKLAMA" in satir_metni:
                    baslik_satiri = i
                    break

            df.columns = [str(c).strip().upper() for c in df.iloc[baslik_satiri]]

            # Verilerin başladığı satırdan itibaren tara
            for index, row in df.iloc[baslik_satiri + 1:].iterrows():
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
                            return (sira_no, "") if resim_no_cek else sira_no
                        except:
                            pass
        except Exception as e:
            print(f"DNA Okuma Hatası: {e}")

        # Eşleşme yoksa (Yepyeni bir özellikse)
        return (None, "") if resim_no_cek else None

    def uret_calistir(self):
        # 1. ARAYÜZDEN VERİLERİ ÇEK
        versiyon = self.versiyon_kutusu.currentText()
        grup = self.grup_kutusu.currentText()
        kutup = self.kutup_kutusu.currentText()
        paket_boyu = self.paket_boyu_input.text().strip().upper()

        # Kullanıcının manuel girdiği revizyonu çekiyoruz ve kutuya (dropdown hafızasına) ekliyoruz
        revizyon = self.revizyon_kutusu.currentText().strip()
        if self.revizyon_kutusu.findText(revizyon) == -1:
            self.revizyon_kutusu.addItem(revizyon)

        ham_ek_aciklama = self.ek_aciklama_input.text().strip().upper()
        ek_aciklama = ham_ek_aciklama.replace("ROTOR PAKET", "").strip()

        # ---------------- 2. EXCEL FORMÜLÜNE GÖRE TASLAK AÇIKLAMA OLUŞTURMA ----------------
        taslak_parcalari = []
        if versiyon == "PREMIUM":
            taslak_parcalari.append(f"{grup}-{kutup}")
            if ek_aciklama:
                taslak_parcalari.append(ek_aciklama)
            taslak_parcalari.append(f"ROTOR PAKET PB:{paket_boyu} MM - R:{revizyon}")
        else:
            taslak_parcalari.append(f"{grup} {kutup}")
            if ek_aciklama:
                taslak_parcalari.append(ek_aciklama)
            taslak_parcalari.append("ROTOR PAKET")

        taslak_aciklama = " ".join(taslak_parcalari)
        taslak_aciklama = " ".join(taslak_aciklama.split())  # Çift boşlukları temizle

        # HATA TESPİTİ İÇİN KONSOLA YAZDIRMA (Excel ile bunu karşılaştır!)
        print("\n" + "="*40)
        print("🔍 YAPAY ZEKAYA GÖNDERİLEN VERİLER")
        print(f"GRUP      : '{grup}'")
        print(f"AÇIKLAMA  : '{taslak_aciklama}'")
        print(f"AÇIKLAMA2 : '{ek_aciklama}'")
        print("="*40 + "\n")

        # ---------------- 3. DNA TESTİ İLE SIRA NO BULMA ----------------
        tam_metin_havuzu = f"{taslak_aciklama} {ek_aciklama}"

        bulunan_eski_sira_no = self.dna_bazli_sira_no_bul(sayfa_adi="ROTOR PAKET", aranan_grup=grup,
                                                          taslak_metin=tam_metin_havuzu)

        from arayuz.mantik.sira_no_dialog import sira_no_secim_yap
        secilen_sira = sira_no_secim_yap(self, bulunan_eski_sira_no, taslak_aciklama)
        if secilen_sira is None:
            return
        elif secilen_sira == -1:
            sira_no_int = self.genel_max_sira_no_bul() + 1
        else:
            sira_no_int = secilen_sira
        sira_no_str = str(sira_no_int)

        aciklama_kod_str = sira_no_str.zfill(4)
        # ---------------- 4. NİHAİ ID OLUŞTURMA ----------------
        id_revizyon = "00" if versiyon == "ESKİ_GÖVDE" else revizyon
        uretilen_id = f"YMM{grup}-RP-{aciklama_kod_str}-{id_revizyon}"

        # ---------------- 5. TABLOYA YAZDIRMA ----------------
        bugunun_tarihi = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm")

        # Tablo Sırası: ["#KOD", "AÇIKLAMA", "GRUP", 'AÇIKLAMA2', 'AÇIKLAMA SIRA NO', 'AÇIKLAMA SIRA KOD', "KAYIT TARİHİ"]
        # DİKKAT: aciklama_kod_str başındaki tek tırnak (') kaldırıldı!
        yeni_kayit = [
            uretilen_id, taslak_aciklama, grup, ek_aciklama, sira_no_str, aciklama_kod_str, bugunun_tarihi
        ]

        mevcut_satir = self.tablo.rowCount()
        self.tablo.insertRow(mevcut_satir)
        for sutun, veri in enumerate(yeni_kayit):
            hucre = QTableWidgetItem(str(veri))
            self.tablo.setItem(mevcut_satir, sutun, hucre)


