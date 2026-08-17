# arayuz/profil_ekrani.py
import json
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QMessageBox, QListWidget, QListWidgetItem, QMenu)
from PyQt6.QtCore import Qt
from arayuz.mantik import oturum
from arayuz.giris_ekrani import KULLANICI_DOSYASI
from arayuz.mantik.json_islemleri import (
    sistem_logu_ekle, KALICI_KLASOR,
    log_aktif_mi, log_aktif_ayarla, log_kayitlarini_sil
)
import os


class KullaniciYonetimPaneli(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sistem Kullanıcıları ve Yetkilendirme")
        self.setFixedSize(650, 450)
        self.setStyleSheet("""
            QListWidget { font-size: 14px; font-weight: bold; border: 1px solid #BDC3C7; border-radius: 5px; padding: 5px;}
            QListWidget::item { padding: 8px; border-bottom: 1px solid #ECF0F1; }
            QListWidget::item:selected { background-color: #3498DB; color: white; border-radius: 4px; }
        """)

        ana_duzen = QHBoxLayout(self)

        # SOL: KULLANICI LİSTESİ
        sol_duzen = QVBoxLayout()
        sol_duzen.addWidget(QLabel("Sistemdeki Kullanıcılar\n(Yetkilendirmek için sağ tıklayın)"))

        self.kullanici_listesi = QListWidget()
        self.kullanici_listesi.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.kullanici_listesi.customContextMenuRequested.connect(self.sag_tik_menusu)
        self.kullanici_listesi.itemClicked.connect(self.detay_goster)

        sol_duzen.addWidget(self.kullanici_listesi)
        ana_duzen.addLayout(sol_duzen, 3)

        # SAĞ: DETAY EKRANI
        sag_duzen = QVBoxLayout()
        self.lbl_detay_baslik = QLabel("Kullanıcı Detayları")
        self.lbl_detay_baslik.setStyleSheet("font-weight: 900; font-size: 16px; color: #2C3E50;")
        sag_duzen.addWidget(self.lbl_detay_baslik)

        self.lbl_detay_ad = QLabel("Ad Soyad: -")
        self.lbl_detay_ad.setStyleSheet("font-size: 14px; margin-top: 10px;")

        self.lbl_detay_kimlik = QLabel("Kimlik No: -")
        self.lbl_detay_kimlik.setStyleSheet("font-size: 14px; margin-top: 5px;")

        self.lbl_detay_yetki = QLabel("Yetki Durumu: -")
        self.lbl_detay_yetki.setStyleSheet("font-size: 14px; margin-top: 5px;")

        sag_duzen.addWidget(self.lbl_detay_ad)
        sag_duzen.addWidget(self.lbl_detay_kimlik)
        sag_duzen.addWidget(self.lbl_detay_yetki)
        sag_duzen.addStretch()

        ana_duzen.addLayout(sag_duzen, 2)
        self.kullanicilari_yukle()

    def kullanicilari_yukle(self):
        self.kullanici_listesi.clear()
        try:
            with open(KULLANICI_DOSYASI, "r", encoding="utf-8") as f:
                self.veriler = json.load(f)

            for kimlik, veri in self.veriler.items():
                ad_soyad = f"{veri['ad']} {veri['soyad']}"
                yetki = veri.get("yetki", "standart")

                if yetki == "as yetkili":
                    ikon = "👑"
                elif yetki == "yetkili":
                    ikon = "⭐"
                else:
                    ikon = "👤"

                item = QListWidgetItem(f"{ikon} {ad_soyad}")
                item.setData(Qt.ItemDataRole.UserRole, kimlik)
                self.kullanici_listesi.addItem(item)
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Kullanıcılar okunamadı:\n{e}")

    def detay_goster(self, item):
        kimlik = item.data(Qt.ItemDataRole.UserRole)
        if kimlik in self.veriler:
            kisi = self.veriler[kimlik]
            self.lbl_detay_ad.setText(f"Ad Soyad: {kisi['ad']} {kisi['soyad']}")
            self.lbl_detay_kimlik.setText(f"Kimlik No: {kimlik}")

            # JSON'da yetkisi bozuksa bile 11111111111'i As Yetkili göster
            yetki = "as yetkili" if kimlik == "11111111111" else kisi.get('yetki', 'standart')

            if yetki == 'as yetkili':
                yetki_metni = "<span style='color:#C0392B; font-weight:bold;'>As Yetkili (Kurucu)</span>"
            elif yetki == 'yetkili':
                yetki_metni = "<span style='color:#27AE60; font-weight:bold;'>Yetkili (Yönetici)</span>"
            else:
                yetki_metni = "<span style='color:#7F8C8D; font-weight:bold;'>Standart Kullanıcı</span>"

            self.lbl_detay_yetki.setText(f"Yetki Durumu: {yetki_metni}")

    def sag_tik_menusu(self, pos):
        item = self.kullanici_listesi.itemAt(pos)
        if not item: return

        kimlik = item.data(Qt.ItemDataRole.UserRole)
        hedef_yetki = "as yetkili" if kimlik == "11111111111" else self.veriler[kimlik].get("yetki", "standart")

        if hedef_yetki == "as yetkili":
            QMessageBox.warning(self, "Erişim Engeli", "As Yetkili hesaba müdahale edilemez!")
            return

        menu = QMenu()
        menu.setStyleSheet("font-size: 14px; padding: 5px;")

        aksiyon_yetkili = None
        aksiyon_standart = None
        aksiyon_sil = None

        if hedef_yetki == "standart":
            aksiyon_yetkili = menu.addAction("⭐ Yetkili Yap")
        elif hedef_yetki == "yetkili":
            aksiyon_standart = menu.addAction("👤 Standart Kullanıcıya Düşür")

        # SADECE AS YETKİLİ SİLEBİLİR (Yetkililer bu butonu göremez bile)
        if oturum.aktif_yetki == "as yetkili":
            menu.addSeparator()
            aksiyon_sil = menu.addAction("🗑️ Kullanıcıyı Sistemden Sil")

        secim = menu.exec(self.kullanici_listesi.viewport().mapToGlobal(pos))

        if secim:
            if secim == aksiyon_yetkili:
                self.yetki_degistir(kimlik, "yetkili")
            elif secim == aksiyon_standart:
                self.yetki_degistir(kimlik, "standart")
            elif secim == aksiyon_sil:
                self.kullanici_sil(kimlik)

    # yetki_degistir fonksiyonunun hemen altına eklenecek yeni fonksiyon:
    def kullanici_sil(self, kimlik):
        cevap = QMessageBox.question(self, "Kullanıcı Sil",
                                     f"Bu kullanıcıyı sistemden KALICI olarak silmek istediğinize emin misiniz?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if cevap == QMessageBox.StandardButton.Yes:
            silinecek_ad = f"{self.veriler[kimlik]['ad']} {self.veriler[kimlik]['soyad']}"  # Silinmeden önce adı alıyoruz
            del self.veriler[kimlik]  # JSON verisinden kimliği uçur
            try:
                with open(KULLANICI_DOSYASI, "w", encoding="utf-8") as f:
                    json.dump(self.veriler, f, indent=4, ensure_ascii=False)

                # --- YENİ LOG EKLENTİSİ ---
                sistem_logu_ekle(
                    f"As Yetkili {oturum.tam_isim()}, '{silinecek_ad}' adlı kullanıcıyı sistemden KALICI olarak sildi.")

                QMessageBox.information(self, "Başarılı", "Kullanıcı sistemden tamamen silindi.")

                self.kullanicilari_yukle()
                self.lbl_detay_ad.setText("Ad Soyad: -")
                self.lbl_detay_kimlik.setText("Kimlik No: -")
                self.lbl_detay_yetki.setText("Yetki Durumu: -")
            except Exception as e:
                QMessageBox.warning(self, "Hata", f"Kullanıcı silinemedi:\n{e}")

    def yetki_degistir(self, kimlik, yeni_yetki):
        # Kendini düşürmeyi engelle
        if kimlik == oturum.aktif_kimlik:
            QMessageBox.warning(self, "Güvenlik Engeli", "Kendi yetkinizi değiştiremezsiniz!")
            return

        self.veriler[kimlik]["yetki"] = yeni_yetki
        try:
            with open(KULLANICI_DOSYASI, "w", encoding="utf-8") as f:
                json.dump(self.veriler, f, indent=4, ensure_ascii=False)

            QMessageBox.information(self, "Başarılı", f"Kullanıcı yetkisi '{yeni_yetki.upper()}' olarak güncellendi.")

            self.kullanicilari_yukle()
            self.lbl_detay_ad.setText("Ad Soyad: -")
            self.lbl_detay_kimlik.setText("Kimlik No: -")
            self.lbl_detay_yetki.setText("Yetki Durumu: -")
            # --- YENİ LOG EKLENTİSİ ---
            islem_goren_ad = f"{self.veriler[kimlik]['ad']} {self.veriler[kimlik]['soyad']}"
            sistem_logu_ekle(
                f"{oturum.tam_isim()} adlı yönetici, '{islem_goren_ad}' kişisinin yetkisini '{yeni_yetki.upper()}' yaptı.")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Yetki güncellenemedi:\n{e}")


class SistemGecmisiPaneli(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AEMOT - Sistem İşlem Geçmişi")
        self.setMinimumSize(850, 600)

        # AEMOT Kurumsal Renkleri (Kırmızı: #C0392B, Koyu Lacivert/Siyah: #1C2833)
        self.setStyleSheet("""
            QDialog {
                background-color: #F4F6F9;
                font-family: 'Segoe UI', Arial;
            }
            QLabel#Baslik {
                color: #C0392B;
                font-size: 22px;
                font-weight: 900;
                letter-spacing: 1px;
            }
            QLabel#Sayac {
                color: #7F8C8D;
                font-size: 14px;
                font-weight: bold;
            }
            QListWidget {
                background-color: #FFFFFF;
                color: #2C3E50;
                font-size: 13px;
                font-weight: 600;
                border: 2px solid #D5D8DC;
                border-radius: 8px;
                padding: 8px;
            }
            QListWidget::item {
                padding: 12px;
                border-bottom: 1px solid #EAEDED;
            }
            QListWidget::item:selected {
                background-color: #FADBD8;
                color: #C0392B;
                border-radius: 5px;
            }
            QPushButton {
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                border: none;
            }
            QPushButton#BtnKapat {
                background-color: #34495E;
                color: white;
            }
            QPushButton#BtnKapat:hover { background-color: #2C3E50; }

            QPushButton#BtnSil {
                background-color: #E74C3C;
                color: white;
            }
            QPushButton#BtnSil:hover { background-color: #C0392B; }

            QPushButton#BtnAktif {
                background-color: #27AE60;
                color: white;
            }
            QPushButton#BtnDeaktif {
                background-color: #F39C12;
                color: white;
            }
        """)

        ana_duzen = QVBoxLayout(self)
        ana_duzen.setSpacing(15)
        ana_duzen.setContentsMargins(25, 25, 25, 25)

        # --- ÜST KISIM (Başlık ve Sayaç) ---
        ust_duzen = QHBoxLayout()
        baslik = QLabel("AEMOT YÖNETİM SİSTEMİ LOGLARI")
        baslik.setObjectName("Baslik")
        ust_duzen.addWidget(baslik)
        ust_duzen.addStretch()

        self.lbl_sayac = QLabel("0 Kayıt")
        self.lbl_sayac.setObjectName("Sayac")
        ust_duzen.addWidget(self.lbl_sayac)
        ana_duzen.addLayout(ust_duzen)

        # --- DURUM BİLGİSİ ---
        self.lbl_log_durum = QLabel("")
        self.lbl_log_durum.setStyleSheet("font-size: 14px; font-weight: bold;")
        ana_duzen.addWidget(self.lbl_log_durum)

        # --- LOG LİSTESİ ---
        self.liste = QListWidget()
        ana_duzen.addWidget(self.liste, stretch=1)

        # --- ALT BUTONLAR ---
        alt_duzen = QHBoxLayout()

        if oturum.aktif_yetki in ["as yetkili", "yetkili"]:
            self.btn_toggle = QPushButton("")
            self.btn_toggle.clicked.connect(self.log_togla)
            alt_duzen.addWidget(self.btn_toggle)

        alt_duzen.addStretch()

        if oturum.aktif_yetki == "as yetkili":
            btn_sil = QPushButton("🗑️ Tüm Kayıtları Sil")
            btn_sil.setObjectName("BtnSil")
            btn_sil.clicked.connect(self.tum_kayitlari_sil)
            alt_duzen.addWidget(btn_sil)

        btn_kapat = QPushButton("Kapat")
        btn_kapat.setObjectName("BtnKapat")
        btn_kapat.clicked.connect(self.accept)
        alt_duzen.addWidget(btn_kapat)

        ana_duzen.addLayout(alt_duzen)

        self._log_durumunu_guncelle()
        self.loglari_yukle()

    def _log_durumunu_guncelle(self):
        aktif = log_aktif_mi()
        if aktif:
            self.lbl_log_durum.setText("🟢 Sistem İzleme AKTİF — İşlemler anlık olarak kaydediliyor.")
            self.lbl_log_durum.setStyleSheet("color: #27AE60; font-size: 13px;")
        else:
            self.lbl_log_durum.setText("🔴 Sistem İzleme DEAKTİF — İşlemler kaydedilmiyor.")
            self.lbl_log_durum.setStyleSheet("color: #E74C3C; font-size: 13px;")

        if hasattr(self, 'btn_toggle'):
            if aktif:
                self.btn_toggle.setText("⏸ Kaydı Durdur")
                self.btn_toggle.setObjectName("BtnDeaktif")
                self.btn_toggle.setStyleSheet("background-color: #F39C12; color: white;")
            else:
                self.btn_toggle.setText("▶ Kaydı Başlat")
                self.btn_toggle.setObjectName("BtnAktif")
                self.btn_toggle.setStyleSheet("background-color: #27AE60; color: white;")

    def loglari_yukle(self):
        self.liste.clear()
        log_dosyasi = os.path.join(KALICI_KLASOR, "sistem_gecmisi.json")
        if os.path.exists(log_dosyasi):
            try:
                with open(log_dosyasi, "r", encoding="utf-8") as f:
                    loglar = json.load(f)
                    if not loglar:
                        self.liste.addItem(" Sistemde henüz kaydedilmiş bir hareket bulunmuyor.")
                        self.lbl_sayac.setText("0 Kayıt")
                        return
                    for log in loglar:
                        item = QListWidgetItem(log)
                        if "SİLDİ" in log.upper() or "KALICI" in log.upper():
                            item.setForeground(Qt.GlobalColor.red)
                        elif "YETKİ" in log.upper() or "YETKILI" in log.upper():
                            item.setForeground(Qt.GlobalColor.blue)
                        self.liste.addItem(item)
                    self.lbl_sayac.setText(f"{len(loglar)} Kayıt")
            except Exception as e:
                self.liste.addItem(f"Log okunamadı: {e}")
        else:
            self.liste.addItem(" Sistemde henüz kaydedilmiş bir hareket bulunmuyor.")
            self.lbl_sayac.setText("0 Kayıt")

    def log_togla(self):
        mevcut = log_aktif_mi()
        yeni_deger = not mevcut

        # Kendi penceremize özel stilize MessageBox (Beyaz kutu hatasını önler)
        cevap = QMessageBox(self)
        cevap.setWindowTitle("Kayıt Durumu Değiştiriliyor")
        cevap.setText("Sistem izleme durumunu değiştirmek istediğinize emin misiniz?")
        cevap.setStyleSheet(
            "QMessageBox { background-color: #F4F6F9; } QLabel { color: #2C3E50; font-size: 13px; font-weight: bold; } QPushButton { background-color: #34495E; color: white; padding: 6px 15px; border-radius: 3px; } QPushButton:hover { background-color: #2C3E50; }")
        cevap.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if cevap.exec() == QMessageBox.StandardButton.Yes:
            log_aktif_ayarla(yeni_deger)
            if yeni_deger:
                sistem_logu_ekle(
                    f"{oturum.tam_isim()} ({oturum.aktif_yetki.upper()}) sistem izlemeyi AKTİF hale getirdi.")
            self._log_durumunu_guncelle()
            self.loglari_yukle()

    def tum_kayitlari_sil(self):
        if oturum.aktif_yetki != "as yetkili":
            return

        cevap = QMessageBox(self)
        cevap.setWindowTitle("Kritik Uyarı")
        cevap.setText(
            "TÜM sistem log kayıtları KALICI OLARAK SİLİNECEK!\n\nBu işlem geri alınamaz. Devam etmek istiyor musunuz?")
        # Kırmızı temalı özel uyarı kutusu
        cevap.setStyleSheet(
            "QMessageBox { background-color: #2C3E50; } QLabel { color: white; font-size: 13px; font-weight: bold; } QPushButton { background-color: #E74C3C; color: white; padding: 6px 15px; border-radius: 3px; font-weight: bold; } QPushButton:hover { background-color: #C0392B; }")
        cevap.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if cevap.exec() == QMessageBox.StandardButton.Yes:
            basarili = log_kayitlarini_sil()
            if basarili:
                sistem_logu_ekle(f"{oturum.tam_isim()} (AS YETKİLİ) tüm sistem log kayıtlarını sildi.")
                self.loglari_yukle()


class ProfilEkrani(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kullanıcı Profili")
        self.setFixedSize(350, 400)

        duzen = QVBoxLayout()
        duzen.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_baslik = QLabel("PROFİL BİLGİLERİ")
        lbl_baslik.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 20px;")
        lbl_baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_isim = QLabel(f"Ad Soyad: {oturum.tam_isim()}")
        lbl_isim.setStyleSheet("font-size: 14px; font-weight: bold; color: #2C3E50;")

        lbl_kimlik = QLabel(f"Kimlik No: {oturum.aktif_kimlik}")
        lbl_kimlik.setStyleSheet("font-size: 14px; font-weight: bold; color: #2C3E50; margin-bottom: 20px;")

        self.yeni_sifre = QLineEdit()
        self.yeni_sifre.setPlaceholderText("Yeni Şifre Belirle")
        self.yeni_sifre.setEchoMode(QLineEdit.EchoMode.Password)
        self.yeni_sifre.setStyleSheet("padding: 8px; font-size: 14px;")

        btn_guncelle = QPushButton("Şifreyi Güncelle")
        btn_guncelle.setStyleSheet("background-color: #E67E22; color: white; font-weight: bold; padding: 10px;")
        btn_guncelle.clicked.connect(self.sifre_guncelle)

        duzen.addWidget(lbl_baslik)
        duzen.addWidget(lbl_isim)
        duzen.addWidget(lbl_kimlik)
        duzen.addWidget(QLabel("Şifre Değiştir:"))
        duzen.addWidget(self.yeni_sifre)
        duzen.addWidget(btn_guncelle)

        # --- YENİ YÖNETİM BUTONU (AS YETKİLİ VE YETKİLİYE ÖZEL) ---
        if oturum.aktif_yetki in ["as yetkili", "yetkili"]:
            duzen.addSpacing(15)
            ayrac = QLabel("───────────")
            ayrac.setAlignment(Qt.AlignmentFlag.AlignCenter)
            duzen.addWidget(ayrac)

            self.btn_yonetici_paneli = QPushButton("👑 Kullanıcı Yönetim Paneli")
            self.btn_yonetici_paneli.setStyleSheet(
                "background-color: #8E44AD; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
            self.btn_yonetici_paneli.clicked.connect(self.yonetici_paneli_ac)
            duzen.addWidget(self.btn_yonetici_paneli)

            # --- YENİ EKLENEN LOG BUTONU ---
            self.btn_loglar = QPushButton("📜 Sistem İşlem Geçmişi")
            self.btn_loglar.setStyleSheet(
                "background-color: #34495E; color: white; font-weight: bold; padding: 10px; border-radius: 4px; margin-top: 5px;")
            self.btn_loglar.clicked.connect(self.log_paneli_ac)
            duzen.addWidget(self.btn_loglar)

        self.setLayout(duzen)

        # ProfilEkrani sınıfının içine şu fonksiyonu da ekle:
    def log_paneli_ac(self):
        panel = SistemGecmisiPaneli(self)
        panel.exec()

    def yonetici_paneli_ac(self):
        panel = KullaniciYonetimPaneli(self)
        panel.exec()

    def sifre_guncelle(self):
        yeni = self.yeni_sifre.text().strip()
        if not yeni:
            QMessageBox.warning(self, "Uyarı", "Yeni şifre boş olamaz!")
            return

        try:
            with open(KULLANICI_DOSYASI, "r", encoding="utf-8") as f:
                kullanicilar = json.load(f)

            if oturum.aktif_kimlik in kullanicilar:
                kullanicilar[oturum.aktif_kimlik]["sifre"] = yeni

                with open(KULLANICI_DOSYASI, "w", encoding="utf-8") as f:
                    json.dump(kullanicilar, f, indent=4, ensure_ascii=False)

                QMessageBox.information(self, "Başarılı", "Şifreniz başarıyla güncellendi!")
                self.accept()
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"İşlem başarısız:\n{e}")