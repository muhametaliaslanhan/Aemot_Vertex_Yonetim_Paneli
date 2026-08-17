import sys
import os
import ctypes
from PyQt6.QtWidgets import QApplication, QMainWindow, QStackedWidget, QSplashScreen, QDialog
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QFont
from PyQt6.QtCore import Qt, QTimer

from arayuz.ana_sayfa import AnaMenuWidget
from arayuz.id_olusturucu_motor import UretimAnaPaneli
from arayuz.urun_agaci import UrunAgaciSekmesi
from arayuz.ayarlar import AyarlarArayuzu
from arayuz.giris_ekrani import GirisPenceresi


def kaynak_yolu(goreceli_yol):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, goreceli_yol)


class AnaUygulama(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VERTEX - AEMOT Yönetim Paneli")

        ikon_tam_yolu = kaynak_yolu("arayuz/varliklar/aemot_ikon.ico")
        self.setWindowIcon(QIcon(ikon_tam_yolu))

        self.resize(1300, 800)
        self.setStyleSheet("background-color: #f0f0f0;")

        # Çıkış yapılıp yapılmadığını anlayan bayrak
        self.cikis_yapildi = False

        self.sayfalar = QStackedWidget()
        self.setCentralWidget(self.sayfalar)

        self.sayfa_ana_menu = AnaMenuWidget(self)
        self.sayfa_motor = UretimAnaPaneli(self)
        self.sayfa_urun_agaci = UrunAgaciSekmesi(self)
        self.sayfa_ayarlar = AyarlarArayuzu(self)

        # GİRİŞ EKRANI BURADAN KALDIRILDI! Artık 0. İndeks direkt Ana Menü
        self.sayfalar.addWidget(self.sayfa_ana_menu)  # İndeks 0
        self.sayfalar.addWidget(self.sayfa_motor)  # İndeks 1
        self.sayfalar.addWidget(self.sayfa_urun_agaci)  # İndeks 2
        self.sayfalar.addWidget(self.sayfa_ayarlar)  # İndeks 3

        self.sayfalar.setCurrentIndex(0)

        # Ana pencere açılır açılmaz profili günceller
        self.sayfa_ana_menu.profil_guncelle()

    def sayfa_degistir(self, indeks):
        self.sayfalar.setCurrentIndex(indeks)


if __name__ == '__main__':
    myappid = 'aemot.vertex.app.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)

    # --- SİSTEM DÖNGÜSÜ BURADA BAŞLIYOR ---
    while True:
        # 1. ÖNCE BAĞIMSIZ GİRİŞ PENCERESİ AÇILIR
        giris = GirisPenceresi()

        # Eğer QDialog olan giriş ekranı başarıyla geçilirse (giris.accept() tetiklenirse)
        if giris.exec() == QDialog.DialogCode.Accepted:

            # 2. BAŞARILI GİRİŞ SONRASI SPLASH (YÜKLENİYOR) EKRANI ÇIKAR
            splash_pixmap = QPixmap(500, 300)
            splash_pixmap.fill(QColor("#FFFFFF"))

            painter = QPainter(splash_pixmap)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)

            font_baslik = QFont("Segoe UI", 36, QFont.Weight.ExtraBold)
            painter.setFont(font_baslik)
            painter.setPen(QColor("#E30613"))
            painter.drawText(splash_pixmap.rect().adjusted(0, -40, 0, 0), Qt.AlignmentFlag.AlignCenter, "AEMOT")

            font_alt = QFont("Segoe UI", 12, QFont.Weight.Bold)
            painter.setFont(font_alt)
            painter.setPen(QColor("#2C3E50"))
            painter.drawText(splash_pixmap.rect().adjusted(0, 30, 0, 0), Qt.AlignmentFlag.AlignCenter,
                             "VERTEX Üretim ve Yönetim Paneli")

            font_yukleniyor = QFont("Segoe UI", 9, QFont.Weight.Medium)
            painter.setFont(font_yukleniyor)
            painter.setPen(QColor("#7F8C8D"))

            # Adamın ismini çekip Splash ekrana basıyoruz
            from arayuz.mantik import oturum

            karsilama = f"Hoş Geldin {oturum.tam_isim()}... Sistem Başlatılıyor"
            painter.drawText(splash_pixmap.rect().adjusted(10, 10, -20, -20),
                             Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight, karsilama)

            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor("#E30613"))
            painter.drawRect(0, 290, 500, 10)

            painter.end()

            splash = QSplashScreen(splash_pixmap, Qt.WindowType.WindowStaysOnTopHint)
            splash.show()
            app.processEvents()

            # 3. ANA UYGULAMAYI ARKA PLANDA YÜKLE
            pencere = AnaUygulama()

            # 1.5 Saniye sonra yükleniyor ekranını kapatıp, asıl pencereyi göster
            QTimer.singleShot(1500, splash.close)
            QTimer.singleShot(1500, pencere.show)

            # Programı çalıştır ve adam kapatana kadar bekle
            app.exec()

            # Eğer adam "Güvenli Çıkış" butonuna basarak programı kapattıysa, bayrak True olur
            if getattr(pencere, "cikis_yapildi", False):
                continue  # Başa sar, giriş ekranını tekrar aç
            else:
                break  # Çarpıdan kapattıysa tamamen çık
        else:
            # Giriş ekranındayken çarpıya basıp kapattıysa direkt çık
            break

    sys.exit()