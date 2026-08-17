import json
import os
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QMessageBox, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPixmap, QCursor
from arayuz.mantik import oturum

# DİKKAT: Kaynak yolunu çekiyoruz ki ikonlar ve logolar EXE'de de çalışsın
from arayuz.mantik.json_islemleri import kaynak_yolu

KULLANICI_DOSYASI = os.path.join(os.getcwd(), "arayuz", "mantik", "kullanicilar.json")

if not os.path.exists(KULLANICI_DOSYASI):
    os.makedirs(os.path.dirname(KULLANICI_DOSYASI), exist_ok=True)
    with open(KULLANICI_DOSYASI, "w", encoding="utf-8") as f:
        # SİSTEM ADMİNİ'ne otomatik 'as yetkili' yetkisi verilir
        ornek_kullanici = {"11111111111": {"ad": "SİSTEM", "soyad": "ADMİNİ", "sifre": "admin", "yetki": "as yetkili"}}
        json.dump(ornek_kullanici, f, indent=4)

def verileri_oku():
    try:
        with open(KULLANICI_DOSYASI, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def verileri_yaz(veriler):
    try:
        with open(KULLANICI_DOSYASI, "w", encoding="utf-8") as f:
            json.dump(veriler, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Kayıt Hatası: {e}")


# --- 1. AYRI KAYIT PENCERESİ ---
class KayitPenceresi(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VERTEX - Yeni Kayıt")
        self.setWindowIcon(QIcon(kaynak_yolu("arayuz/varliklar/aemot_ikon.ico")))  # Özel AEMOT İkonu
        self.setFixedSize(350, 450)
        self.setStyleSheet("background-color: #FFFFFF;")

        duzen = QVBoxLayout()
        duzen.setContentsMargins(30, 30, 30, 30)
        duzen.setSpacing(15)

        baslik = QLabel("Sisteme Kayıt Ol")
        baslik.setStyleSheet("font-size: 20px; font-weight: bold; color: #2C3E50;")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)

        style_input = "padding: 12px; font-size: 14px; border: 1px solid #BDC3C7; border-radius: 5px; background-color: #F8F9F9;"

        self.kayit_ad = QLineEdit()
        self.kayit_ad.setPlaceholderText("Adınız")
        self.kayit_ad.setStyleSheet(style_input)

        self.kayit_soyad = QLineEdit()
        self.kayit_soyad.setPlaceholderText("Soyadınız")
        self.kayit_soyad.setStyleSheet(style_input)

        self.kayit_kimlik = QLineEdit()
        self.kayit_kimlik.setPlaceholderText("Kimlik Numaranız")
        self.kayit_kimlik.setMaxLength(11)  # KULLANICI 11'DEN FAZLA YAZAMAZ
        self.kayit_kimlik.setStyleSheet(style_input)

        self.kayit_sifre = QLineEdit()
        self.kayit_sifre.setPlaceholderText("Şifre Belirleyin")
        self.kayit_sifre.setEchoMode(QLineEdit.EchoMode.Password)
        self.kayit_sifre.setStyleSheet(style_input)

        btn_kayit = QPushButton("KAYDI TAMAMLA")
        btn_kayit.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_kayit.setStyleSheet("""
            QPushButton { background-color: #27AE60; color: white; font-weight: bold; padding: 12px; font-size: 14px; border-radius: 5px; }
            QPushButton:hover { background-color: #2ECC71; }
        """)
        btn_kayit.clicked.connect(self.kayit_ol)

        duzen.addWidget(baslik)
        duzen.addSpacing(10)
        duzen.addWidget(self.kayit_ad)
        duzen.addWidget(self.kayit_soyad)
        duzen.addWidget(self.kayit_kimlik)
        duzen.addWidget(self.kayit_sifre)
        duzen.addSpacing(15)
        duzen.addWidget(btn_kayit)
        duzen.addStretch()

        self.setLayout(duzen)

    def kayit_ol(self):
        ad = self.kayit_ad.text().strip().upper()
        soyad = self.kayit_soyad.text().strip().upper()
        kimlik = self.kayit_kimlik.text().strip()
        sifre = self.kayit_sifre.text().strip()

        if not ad or not soyad or not kimlik or not sifre:
            QMessageBox.warning(self, "Uyarı", "Lütfen tüm alanları doldurun!")
            return

        kullanicilar = verileri_oku()
        if kimlik in kullanicilar:
            QMessageBox.warning(self, "Hata", "Bu kimlik numarası ile zaten kayıtlı bir kullanıcı var!")
            return

        # YENİ KAYIT OLAN HERKES OTOMATİK STANDART KULLANICI OLUR
        kullanicilar[kimlik] = {"ad": ad, "soyad": soyad, "sifre": sifre, "yetki": "standart"}
        verileri_yaz(kullanicilar)

        QMessageBox.information(self, "Başarılı", "Kayıt tamamlandı! Şimdi giriş yapabilirsiniz.")
        self.accept()

# --- 2. ANA GİRİŞ PENCERESİ ---
class GirisPenceresi(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VERTEX - Kullanıcı Girişi")
        self.setWindowIcon(QIcon(kaynak_yolu("arayuz/varliklar/aemot_ikon.ico")))  # Özel AEMOT İkonu
        self.setFixedSize(400, 500)
        self.setStyleSheet("background-color: #FFFFFF;")

        ana_duzen = QVBoxLayout()
        ana_duzen.setContentsMargins(40, 40, 40, 40)
        ana_duzen.setSpacing(15)
        ana_duzen.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- AEMOT LOGOSU ---
        self.logo_etiketi = QLabel()
        self.logo_etiketi.setPixmap(QPixmap(kaynak_yolu("arayuz/varliklar/aemot_logo_gorseli.png")))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(220, 60)
        self.logo_etiketi.setAlignment(Qt.AlignmentFlag.AlignCenter)

        logo_duzen = QHBoxLayout()
        logo_duzen.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_duzen.addWidget(self.logo_etiketi)

        baslik = QLabel("VERTEX YÖNETİM PANELİ")
        baslik.setStyleSheet("font-size: 16px; font-weight: bold; color: #7F8C8D; margin-bottom: 10px;")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- INPUTLAR ---
        style_input = "padding: 14px; font-size: 14px; border: 1px solid #BDC3C7; border-radius: 6px; background-color: #F8F9F9;"

        self.giris_kimlik = QLineEdit()
        self.giris_kimlik.setPlaceholderText("Kimlik Numaranız")
        self.giris_kimlik.setMaxLength(11)  # KULLANICI 11'DEN FAZLA YAZAMAZ
        self.giris_kimlik.setStyleSheet(style_input)

        self.giris_sifre = QLineEdit()
        self.giris_sifre.setPlaceholderText("Şifreniz")
        self.giris_sifre.setEchoMode(QLineEdit.EchoMode.Password)
        self.giris_sifre.setStyleSheet(style_input)

        # --- GİRİŞ BUTONU ---
        btn_giris = QPushButton("SİSTEME GİRİŞ YAP")
        btn_giris.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_giris.setStyleSheet("""
            QPushButton { background-color: #E30613; color: white; font-weight: bold; padding: 14px; font-size: 15px; border-radius: 6px; }
            QPushButton:hover { background-color: #C0392B; }
        """)
        btn_giris.clicked.connect(self.giris_yap)

        # --- AYRAÇ VE KAYIT OL LİNKİ ---
        ayrac = QLabel("────────  veya  ────────")
        ayrac.setStyleSheet(
            "color: #BDC3C7; font-size: 12px; font-weight: bold; margin-top: 15px; margin-bottom: 15px;")
        ayrac.setAlignment(Qt.AlignmentFlag.AlignCenter)

        btn_kayit_ekrani = QPushButton("Hesabınız yok mu? Yeni Kayıt Ol")
        btn_kayit_ekrani.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_kayit_ekrani.setStyleSheet("""
            QPushButton { background-color: transparent; color: #2980B9; font-weight: bold; font-size: 13px; text-decoration: underline; border: none; }
            QPushButton:hover { color: #1ABC9C; }
        """)
        btn_kayit_ekrani.clicked.connect(self.kayit_ekranini_ac)

        # DÜZENE EKLE
        ana_duzen.addLayout(logo_duzen)
        ana_duzen.addWidget(baslik)
        ana_duzen.addSpacing(10)
        ana_duzen.addWidget(self.giris_kimlik)
        ana_duzen.addWidget(self.giris_sifre)
        ana_duzen.addSpacing(10)
        ana_duzen.addWidget(btn_giris)
        ana_duzen.addWidget(ayrac)
        ana_duzen.addWidget(btn_kayit_ekrani)
        ana_duzen.addStretch()

        self.setLayout(ana_duzen)

    def giris_yap(self):
        kimlik = self.giris_kimlik.text().strip()
        sifre = self.giris_sifre.text().strip()

        kullanicilar = verileri_oku()

        if kimlik in kullanicilar and kullanicilar[kimlik]["sifre"] == sifre:
            oturum.aktif_kimlik = kimlik
            oturum.aktif_ad = kullanicilar[kimlik]["ad"]
            oturum.aktif_soyad = kullanicilar[kimlik]["soyad"]

            # Kurucu kimlik JSON'da 'as yetkili' değilse zorla düzelt ve KALICI kaydet!
            if kimlik == "11111111111" and kullanicilar[kimlik].get("yetki") != "as yetkili":
                kullanicilar[kimlik]["yetki"] = "as yetkili"
                verileri_yaz(kullanicilar)  # json'a yazıyoruz ki liste güncellensin

            oturum.aktif_yetki = kullanicilar[kimlik].get("yetki", "standart")
            self.accept()
        else:
            QMessageBox.warning(self, "Hata", "Kimlik numarası veya şifre hatalı!")

    def kayit_ekranini_ac(self):
        # Butona basılınca diğer ufak pencereyi çağırıyoruz
        pencere_kayit = KayitPenceresi()
        pencere_kayit.exec()