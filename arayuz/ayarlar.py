from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QLineEdit, QGroupBox, QFileDialog, QMessageBox, QDialog, QFormLayout, QComboBox, QToolTip)

from arayuz.mantik.json_islemleri import (katalog_oku, katalog_yaz, standart_flans_oku,
                                          standart_flans_yaz, ayar_excel_yolunu_getir, ayar_excel_yolunu_kaydet, kaynak_yolu,
                                          urun_agaci_excel_yolunu_getir, urun_agaci_excel_yolunu_kaydet)

from arayuz.mantik.motor_verileri import turkce_karakter_temizleme, motor_ozellikleri
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QCursor
from arayuz.mantik import oturum
import pandas as pd
import json
import os



# --- 1. KATALOG EKLEME PENCERESİ (Buraya taşıdık) ---
class KatalogEklePenceresi(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Katalog Veri Ekleme")
        self.setFixedSize(500, 350)
        self.setModal(True)

        ana_duzen = QVBoxLayout()

        baslik = QLabel("AEMOT ÖZEL KW VE DEVİR EKLEME")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)
        baslik.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px; color: #2C3E50;")
        ana_duzen.addWidget(baslik)

        form_duzen = QFormLayout()

        self.combo_M = QComboBox()
        self.combo_govde = QComboBox()
        self.combo_kutup = QComboBox()
        self.combo_paket = QComboBox()

        self.combo_M.addItems(motor_ozellikleri.ceviri_katologu["M"].keys())
        self.combo_govde.addItems(motor_ozellikleri.ceviri_katologu["govde_boyu"].keys())
        self.combo_kutup.addItems(motor_ozellikleri.ceviri_katologu["kutup"].keys())
        self.combo_paket.addItems(motor_ozellikleri.ceviri_katologu["paket_boyu"].keys())

        self.input_kw = QLineEdit()
        self.input_kw.setPlaceholderText("Örn: 0.18 veya OZEL")
        self.input_devir = QLineEdit()
        self.input_devir.setPlaceholderText("Örn: 1400 veya OZEL")

        form_duzen.addRow("M (Gövde Tipi):", self.combo_M)
        form_duzen.addRow("Gövde Boyu:", self.combo_govde)
        form_duzen.addRow("Kutup:", self.combo_kutup)
        form_duzen.addRow("Paket Boyu:", self.combo_paket)
        form_duzen.addRow("KW Değeri:", self.input_kw)
        form_duzen.addRow("Devir (d/dk):", self.input_devir)

        ana_duzen.addLayout(form_duzen)

        buton_duzen = QHBoxLayout()
        self.btn_iptal = QPushButton("İptal")
        self.btn_iptal.setStyleSheet("padding: 8px; font-weight: bold; background-color: #E74C3C; color: white;")
        self.btn_kaydet = QPushButton("Kaydet")
        self.btn_kaydet.setStyleSheet("padding: 8px; font-weight: bold; background-color: #2ECC71; color: white;")

        buton_duzen.addWidget(self.btn_iptal)
        buton_duzen.addWidget(self.btn_kaydet)

        ana_duzen.addLayout(buton_duzen)
        self.setLayout(ana_duzen)

        self.btn_iptal.clicked.connect(self.close)
        self.btn_kaydet.clicked.connect(self.kaydet_fonksiyonu)

    def kaydet_fonksiyonu(self):
        ham_kw = self.input_kw.text().strip()
        ham_devir = self.input_devir.text().strip()

        if not ham_kw or not ham_devir:
            QMessageBox.warning(self, "HATA", "KW ve Devir değerleri boş bırakılamaz!")
            return

        temiz_kw = turkce_karakter_temizleme(ham_kw)
        temiz_devir = turkce_karakter_temizleme(ham_devir)
        m = self.combo_M.currentText()
        govde = self.combo_govde.currentText()
        kutup = self.combo_kutup.currentText()
        paket = self.combo_paket.currentText()
        sifre = f"{m}{govde}{kutup}{paket}"
        katalog = katalog_oku()

        if sifre in katalog:
            cevap = QMessageBox.question(self, "KAYIT ZATEN MEVCUT",
                                         f"Bu motor zaten var! Güncellemek ister misin?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        try:
            katalog[sifre] = {"kw": temiz_kw, "devir": temiz_devir}
            katalog_yaz(katalog)
            QMessageBox.information(self, "BAŞARILI", "Değerler JSON Veritabanına Eklendi!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "HATA", f"Kaydederken sorun oluştu:\n{e}")


# --- 2. STANDART FLANŞ EKLEME PENCERESİ (Buraya taşıdık) ---
class StandartFlansEklePenceresi(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Standart Flanş Ekleme")
        self.setFixedSize(400, 220)
        self.setModal(True)

        ana_duzen = QVBoxLayout(self)

        baslik = QLabel("STANDART FLANŞ ÖLÇÜSÜ EKLEME")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)
        baslik.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px; color: #2C3E50;")
        ana_duzen.addWidget(baslik)

        form_duzen = QFormLayout()

        self.combo_m = QComboBox()
        self.combo_m.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630"])
        self.combo_yapi = QComboBox()
        self.combo_yapi.addItems(["B3", "B5", "B9", "B14", "B34"])
        self.combo_flans = QComboBox()
        self.combo_flans.addItems(
            ["YOK", "C90", "C105", "C120", "C140", "C160", "C200", "C250", "A140", "A160", "A200", "A250", "A300",
             "A350", "A400", "A450", "A550", "A660", "A800", "A1000", "A1150", "A1370"])

        form_duzen.addRow("M (Gövde Tipi):", self.combo_m)
        form_duzen.addRow("Yapı Şekli:", self.combo_yapi)
        form_duzen.addRow("Flanş Ölçüsü:", self.combo_flans)
        ana_duzen.addLayout(form_duzen)

        buton_duzen = QHBoxLayout()
        btn_iptal = QPushButton("İptal")
        btn_iptal.setStyleSheet("padding: 8px; font-weight: bold; background-color: #E74C3C; color: white;")
        btn_kaydet = QPushButton("Kaydet")
        btn_kaydet.setStyleSheet("padding: 8px; font-weight: bold; background-color: #2ECC71; color: white;")
        btn_iptal.clicked.connect(self.close)
        btn_kaydet.clicked.connect(self.kaydet_fonksiyonu)
        buton_duzen.addWidget(btn_iptal)
        buton_duzen.addWidget(btn_kaydet)
        ana_duzen.addLayout(buton_duzen)

    def kaydet_fonksiyonu(self):
        m = self.combo_m.currentText().strip()
        yapi = self.combo_yapi.currentText().strip()
        flans = self.combo_flans.currentText().strip()
        sifre = f"{m}{yapi}"
        mevcut_veriler = standart_flans_oku()

        if sifre in mevcut_veriler:
            cevap = QMessageBox.question(self, "KAYIT ZATEN MEVCUT", "Bu gövde JSON'da var. Güncellemek ister misin?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if cevap == QMessageBox.StandardButton.No: return

        yeni_kayit = {"M": m, "Yapi Sekli": yapi, "flans olcusu": flans, "standart bilgi": "STANDART FLANS"}
        mevcut_veriler[sifre] = yeni_kayit
        try:
            standart_flans_yaz(mevcut_veriler)
            QMessageBox.information(self, "BAŞARILI", "Standart flanş ölçüsü JSON'a eklendi!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "HATA", f"Kaydedilirken hata oluştu:\n{e}")


class YasakliGovdeEkleDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⛔ Yasaklı Gövde Kuralı Ekle")
        self.setMinimumWidth(450)
        self.setStyleSheet("""
            QDialog { background-color: #F8F9F9; font-family: 'Segoe UI', Arial; }
            QComboBox { padding: 8px; border: 1px solid #BDC3C7; border-radius: 4px; background: white; font-weight: bold; }
            QLabel { font-weight: bold; color: #2C3E50; }
        """)

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.versiyon_kutusu = QComboBox()
        self.versiyon_kutusu.addItems(["PREMIUM", "ESKİ_GÖVDE"])

        self.grup_kutusu = QComboBox()
        self.grup_kutusu.addItems(
            ["063", "071", "080", "090", "100", "112", "132", "160", "180", "200", "225", "250", "280", "315", "355",
             "400", "450", "500", "630"])

        self.malzeme_kutusu = QComboBox()
        self.malzeme_kutusu.addItems(["ALUMINYUM", "PIK", "CELIK"])

        self.ayak_kutusu = QComboBox()
        self.ayak_kutusu.addItems(["AYAKLI", "AYAKSIZ"])

        self.islenmis_kutusu = QComboBox()
        self.islenmis_kutusu.addItems(["ISLENMIS", "ISLENMEMIS"])

        self.revizyon_kutusu = QComboBox()
        self.revizyon_kutusu.addItems(["01", "02", "03", "04", "05"])

        form_layout.addRow("Versiyon:", self.versiyon_kutusu)
        form_layout.addRow("Grup:", self.grup_kutusu)
        form_layout.addRow("Malzeme:", self.malzeme_kutusu)
        form_layout.addRow("Ayaklı/Ayaksız:", self.ayak_kutusu)
        form_layout.addRow("İşlenmiş Durumu:", self.islenmis_kutusu)
        form_layout.addRow("Revizyon No:", self.revizyon_kutusu)

        layout.addLayout(form_layout)

        self.btn_kaydet = QPushButton("YASAK KURALINI SİSTEME EKLE")
        self.btn_kaydet.setStyleSheet(
            "background-color: #E74C3C; color: white; font-weight: bold; padding: 12px; border-radius: 6px; font-size: 13px;")
        self.btn_kaydet.clicked.connect(self.yasagi_kaydet)
        layout.addWidget(self.btn_kaydet)

    def yasagi_kaydet(self):
        versiyon = self.versiyon_kutusu.currentText().strip()
        grup = self.grup_kutusu.currentText().strip()
        malzeme = self.malzeme_kutusu.currentText().strip()
        ayak = self.ayak_kutusu.currentText().strip()
        islenmis = self.islenmis_kutusu.currentText().strip()
        revizyon = self.revizyon_kutusu.currentText().strip()

        # Gövde formülüne göre Açıklama (Kural) Cümlesini Oluştur
        rev_metni = f"R:{revizyon}" if versiyon != "ESKİ_GÖVDE" else ""
        kontrol_metni = f"{grup} {malzeme} {ayak} {islenmis} {rev_metni}".upper().strip()
        kontrol_metni = " ".join(kontrol_metni.split())  # Fazla boşlukları yutar

        # SENİN DOSYA YOLLARIN
        # DİNAMİK DOSYA YOLLARI (json_islemleri'ndeki kaynak_yolu'nu kullanarak)
        excel_yolu = kaynak_yolu('arayuz/mantik/katalog_veri_json/govde_olusturulamayacak kodlar.xlsx')
        json_yolu = kaynak_yolu('arayuz/mantik/katalog_veri_json/govde_yasaklari.json')

        cevap = QMessageBox.question(self, "Emin misin?",
                                     f"Şu kombinasyon üretim yasaklılar listesine eklenecek:\n\n{kontrol_metni}\n\nOnaylıyor musun?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if cevap == QMessageBox.StandardButton.No:
            return

        try:
            # 1. EXCEL'E EKLE
            df = pd.read_excel(excel_yolu)
            df.columns = df.columns.str.strip().str.upper()  # Güvenlik için başlıkları temizle

            # Kural zaten varsa uyar
            if kontrol_metni in df['ACIKLAMA'].str.strip().str.upper().values:
                QMessageBox.warning(self, "Zaten Var!", "Bu kural Excel'de zaten mevcut!")
                return

            # Gönderdiğin resimdeki sütun başlıklarına tam uygun veri yapısı
            yeni_kayit = pd.DataFrame([{
                "GRUP": grup,
                "MALZEME": malzeme,
                "AYAKLI\AYAKSIZ": ayak,
                "ISLENMIS\ISLENMEMIS": islenmis,
                "REVIZYON_NO": rev_metni,
                "ACIKLAMA": kontrol_metni,
                "SONUC": "KOD OLUŞTURULAMAZ."
            }])

            df = pd.concat([df, yeni_kayit], ignore_index=True)
            df.to_excel(excel_yolu, index=False)

            # 2. JSON'A EKLE (Sistemin anında tanıması için)
            if os.path.exists(json_yolu):
                with open(json_yolu, 'r', encoding='utf-8') as f:
                    yasaklar = json.load(f)
            else:
                yasaklar = {}

            yasaklar[kontrol_metni] = True

            with open(json_yolu, 'w', encoding='utf-8') as f:
                json.dump(yasaklar, f, indent=4)

            QMessageBox.information(self, "İşlem Başarılı",
                                    f"Yasak kuralı Excel'e ve sisteme kusursuz işlendi!\n\nEngellenen Gövde: {kontrol_metni}")
            self.accept()  # Pencereyi kapat

        except PermissionError:
            QMessageBox.critical(self, "Excel Açık!",
                                 "Excel dosyası şu an açık olduğu için işlem yapılamadı. Dosyayı kapatıp tekrar deneyin.")
        except Exception as e:
            QMessageBox.critical(self, "Sistem Hatası", f"Beklenmeyen bir hata oluştu:\n{e}")



# --- 3. ANA AYARLAR SAYFASI ---
class AyarlarArayuzu(QWidget):
    def __init__(self, ana_pencere):
        super().__init__()
        self.ana_pencere = ana_pencere

        # Tooltip (Baloncuk) stilini çizgi roman tarzına benzetiyoruz
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', Arial; font-size: 14px; }
            QPushButton { font-weight: bold; border-radius: 4px; padding: 10px; }
            QLineEdit { padding: 8px; border: 1px solid #BDC3C7; border-radius: 4px; }
            QGroupBox { font-weight: bold; border: 2px solid #2C3E50; border-radius: 6px; margin-top: 15px; padding-top: 15px; }
            QToolTip { 
                background-color: #F1C40F; 
                color: #2C3E50; 
                border: 2px solid #E67E22; 
                border-radius: 8px; 
                font-weight: bold; 
                padding: 5px; 
            }
        """)

        ana_duzen = QVBoxLayout()
        self.setLayout(ana_duzen)

        baslik = QLabel("SİSTEM AYARLARI")
        baslik.setStyleSheet("font-size: 20px; font-weight: bold; color: #2C3E50; margin-bottom: 20px;")
        ana_duzen.addWidget(baslik)

        # ---------------- EXCEL YOLU VE UYARI SİSTEMİ ----------------
        grup_excel = QGroupBox("Veritabanı (Excel) Yolu Ayarı")
        duzen_excel_ana = QVBoxLayout()  # Üstte kutu, altta uyarı yazısı için Dikey Düzen

        duzen_satir = QHBoxLayout()
        self.excel_yolu_kutusu = QLineEdit()

        # Güncel yolu getir ve kutuya yaz
        guncel_yol = ayar_excel_yolunu_getir()
        self.excel_yolu_kutusu.setText(guncel_yol)

        # Yol değiştiğinde otomatik kontrol et
        self.excel_yolu_kutusu.textChanged.connect(self.dosya_yolu_kontrol_et)

        self.btn_gozat = QPushButton("Gözat")
        self.btn_gozat.setStyleSheet("background-color: #F39C12; color: white;")
        self.btn_gozat.clicked.connect(self.excel_sec)

        self.btn_kaydet = QPushButton("Yolu Kaydet")
        self.btn_kaydet.setStyleSheet("background-color: #27AE60; color: white;")
        self.btn_kaydet.clicked.connect(self.excel_yolu_kaydet)

        # Çizgi roman balonu için ünlem butonu (yanında duracak)
        self.ikon_uyari = QLabel("⚠️")
        self.ikon_uyari.setStyleSheet("font-size: 20px;")
        self.ikon_uyari.setToolTip("Hey! Veritabanı Bulunamadı!\nDosya silinmiş veya yeri değişmiş olabilir!")
        self.ikon_uyari.setVisible(False)

        duzen_satir.addWidget(self.excel_yolu_kutusu)
        duzen_satir.addWidget(self.ikon_uyari)
        duzen_satir.addWidget(self.btn_gozat)
        duzen_satir.addWidget(self.btn_kaydet)

        # Kutunun altına gelecek kırmızı yazı
        self.etiket_alt_uyari = QLabel(
            "Hata: Belirtilen yolda Excel dosyası bulunamadı! Lütfen geçerli bir dosya seçin.")
        self.etiket_alt_uyari.setStyleSheet(
            "color: #E74C3C; font-weight: bold; font-style: italic; font-size: 12px; margin-top: 5px;")
        self.etiket_alt_uyari.setVisible(False)

        duzen_excel_ana.addLayout(duzen_satir)
        duzen_excel_ana.addWidget(self.etiket_alt_uyari)

        grup_excel.setLayout(duzen_excel_ana)
        ana_duzen.addWidget(grup_excel)

        # Sayfa açılır açılmaz bir kez kontrol et
        self.dosya_yolu_kontrol_et()

        # ---------------- ÜRÜN AĞACI EXCEL YOLU AYARI ----------------
        grup_urun_agaci = QGroupBox("Ürün Ağacı Kayıt Exceli")
        duzen_urun_agaci_ana = QVBoxLayout()

        duzen_urun_agaci_satir = QHBoxLayout()
        self.urun_agaci_excel_yolu_kutusu = QLineEdit()

        # Güncel yolu getir
        self.urun_agaci_excel_yolu_kutusu.setText(urun_agaci_excel_yolunu_getir())

        self.btn_urun_agaci_gozat = QPushButton("Gözat")
        self.btn_urun_agaci_gozat.setStyleSheet("background-color: #F39C12; color: white;")
        self.btn_urun_agaci_gozat.clicked.connect(self.urun_agaci_excel_sec)

        self.btn_urun_agaci_kaydet = QPushButton("Yolu Kaydet")
        self.btn_urun_agaci_kaydet.setStyleSheet("background-color: #27AE60; color: white;")
        self.btn_urun_agaci_kaydet.clicked.connect(self.urun_agaci_excel_yolu_kaydet)

        duzen_urun_agaci_satir.addWidget(self.urun_agaci_excel_yolu_kutusu)
        duzen_urun_agaci_satir.addWidget(self.btn_urun_agaci_gozat)
        duzen_urun_agaci_satir.addWidget(self.btn_urun_agaci_kaydet)

        duzen_urun_agaci_ana.addLayout(duzen_urun_agaci_satir)
        grup_urun_agaci.setLayout(duzen_urun_agaci_ana)
        ana_duzen.addWidget(grup_urun_agaci)

        # ---------------- DİĞER BUTONLAR ----------------
        grup_eklemeler = QGroupBox("Katalog ve Flanş Yönetimi")
        duzen_eklemeler = QVBoxLayout()

        self.btn_katalog_ekle = QPushButton("Kataloğa Yeni Özel KW/Devir Ekle")
        self.btn_flans_ekle = QPushButton("Standart Flanş Ölçüsü Ekle")
        self.btn_yasak_ekle = QPushButton("Yasaklı Gövde Kuralı Ekle")  # YENİ BUTON

        self.btn_katalog_ekle.setStyleSheet(
            "background-color: #000080; color: white; font-weight: bold; padding: 10px;")
        self.btn_flans_ekle.setStyleSheet("background-color: #1E90FF; color: white; font-weight: bold; padding: 10px;")
        self.btn_yasak_ekle.setStyleSheet(
            "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px;")  # YENİ BUTON STİLİ

        self.btn_katalog_ekle.clicked.connect(self.katalog_penceresini_ac)
        self.btn_flans_ekle.clicked.connect(self.flans_penceresini_ac)
        self.btn_yasak_ekle.clicked.connect(self.yasakli_govde_penceresi_ac)  # YENİ BUTON BAĞLANTISI

        duzen_eklemeler.addWidget(self.btn_katalog_ekle)
        duzen_eklemeler.addWidget(self.btn_flans_ekle)
        duzen_eklemeler.addWidget(self.btn_yasak_ekle)  # YENİ BUTONU EKRANA EKLE
        grup_eklemeler.setLayout(duzen_eklemeler)

        # --- YETKİLENDİRME KONTROLÜ (GÜVENLİK DUVARI) ---
        # Sadece 'as yetkili' ve 'yetkili' olanlar bu grubu görür, standart ise HİÇ GÖRMEZ
        if oturum.aktif_yetki not in ["as yetkili", "yetkili"]:
            grup_eklemeler.setVisible(False)
        else:
            ana_duzen.addWidget(grup_eklemeler)

        ana_duzen.addStretch()

        btn_geri = QPushButton("Ana Menüye Dön")
        btn_geri.setStyleSheet("background-color: #E74C3C; color: white; padding: 15px;")
        btn_geri.clicked.connect(lambda: self.ana_pencere.sayfa_degistir(0))
        ana_duzen.addWidget(btn_geri)

    # --- YENİ EKLENEN KONTROL SİSTEMİ ---
    def dosya_yolu_kontrol_et(self):
        mevcut_yol = self.excel_yolu_kutusu.text().strip()
        if not os.path.exists(mevcut_yol) or not (mevcut_yol.endswith(".xlsx") or mevcut_yol.endswith(".xls")):
            # Hata var, uyarıları aç ve balonu patlat
            self.excel_yolu_kutusu.setStyleSheet(
                "padding: 8px; border: 2px solid #E74C3C; border-radius: 4px; background-color: #FDEDEC;")
            self.ikon_uyari.setVisible(True)
            self.etiket_alt_uyari.setVisible(True)
        else:
            # Hata yok, her şey normal
            self.excel_yolu_kutusu.setStyleSheet(
                "padding: 8px; border: 1px solid #BDC3C7; border-radius: 4px; background: white;")
            self.ikon_uyari.setVisible(False)
            self.etiket_alt_uyari.setVisible(False)

    def excel_sec(self):
        dosya_yolu, _ = QFileDialog.getOpenFileName(self, "Excel Veritabanını Seç", "",
                                                    "Excel Dosyaları (*.xlsx *.xls)")
        if dosya_yolu:
            self.excel_yolu_kutusu.setText(dosya_yolu)

    def excel_yolu_kaydet(self):
        secilen_yol = self.excel_yolu_kutusu.text().strip()

        if not os.path.exists(secilen_yol):
            QMessageBox.warning(self, "HATA",
                                "Seçilen Excel dosyası bilgisayarda bulunamadı!\nLütfen doğru bir yol seçin.")
            return

        if secilen_yol.endswith(".xlsx") or secilen_yol.endswith(".xls"):
            basarili = ayar_excel_yolunu_kaydet(secilen_yol)
            if basarili:
                QMessageBox.information(self, "BAŞARILI", "Yeni Excel veritabanı yolu KALICI olarak kaydedildi!")
            else:
                QMessageBox.critical(self, "HATA", "Yol ayar dosyasına yazılamadı.")
        else:
            QMessageBox.warning(self, "HATA", "Lütfen geçerli bir Excel dosyası (.xlsx veya .xls) seçin!")

    def urun_agaci_excel_sec(self):
        dosya_yolu, _ = QFileDialog.getSaveFileName(self, "Ürün Ağacı Kayıt Excelini Seç/Oluştur", "",
                                                    "Excel Dosyaları (*.xlsx *.xls)")
        if dosya_yolu:
            self.urun_agaci_excel_yolu_kutusu.setText(dosya_yolu)

    def urun_agaci_excel_yolu_kaydet(self):
        secilen_yol = self.urun_agaci_excel_yolu_kutusu.text().strip()
        if not secilen_yol:
            QMessageBox.warning(self, "HATA", "Yol boş olamaz!")
            return

        if secilen_yol.endswith(".xlsx") or secilen_yol.endswith(".xls"):
            basarili = urun_agaci_excel_yolunu_kaydet(secilen_yol)
            if basarili:
                QMessageBox.information(self, "BAŞARILI", "Ürün Ağacı Excel yolu başarıyla kaydedildi!")
            else:
                QMessageBox.critical(self, "HATA", "Yol ayar dosyasına yazılamadı.")
        else:
            QMessageBox.warning(self, "HATA", "Lütfen geçerli bir Excel dosyası (.xlsx veya .xls) yolu belirtin!")

    def katalog_penceresini_ac(self):
        pencere = KatalogEklePenceresi(self)
        pencere.exec()

    def flans_penceresini_ac(self):
        pencere = StandartFlansEklePenceresi(self)
        pencere.exec()

    def yasakli_govde_penceresi_ac(self):
        dialog = YasakliGovdeEkleDialog(self)
        dialog.exec()