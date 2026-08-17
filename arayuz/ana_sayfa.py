from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QSpacerItem, QSizePolicy, QFrame
)
from PyQt6.QtCore import Qt, QTimer, QDateTime
from PyQt6.QtGui import QPixmap, QCursor
from arayuz.mantik import oturum
from arayuz.profil_ekrani import ProfilEkrani
from PyQt6.QtWidgets import QPushButton, QHBoxLayout, QMessageBox
from arayuz.mantik.json_islemleri import kaynak_yolu


class AnaMenuWidget(QWidget):
    def __init__(self, ana_pencere):
        super().__init__()
        self.ana_pencere = ana_pencere

        # --- ANA STİL DOSYASI (QSS) ---
        self.setStyleSheet("""
            QWidget { 
                font-family: 'Segoe UI', Arial, sans-serif; 
                background-color: #ECF0F1; 
            }
            #KartFrame { 
                background-color: white; 
                border-radius: 12px; 
                border: 1px solid #BDC3C7; 
            }
            QPushButton { 
                border-radius: 6px; 
                color: white; 
                font-weight: bold; 
                padding: 15px; 
                font-size: 16px; 
                text-align: left;
                padding-left: 20px;
            }
            /* Özel Buton Renkleri ve Hover (Üstüne gelince) Efektleri */
            QPushButton#btn1 { background-color: #2980B9; border: 1px solid #1F618D; }
            QPushButton#btn1:hover { background-color: #3498DB; }

            QPushButton#btn2 { background-color: #D35400; border: 1px solid #A04000; }
            QPushButton#btn2:hover { background-color: #E67E22; }

            QPushButton#btn3 { background-color: #7F8C8D; border: 1px solid #616A6B; }
            QPushButton#btn3:hover { background-color: #95A5A6; }
        """)

        ana_duzen = QVBoxLayout(self)

        # --- SOL ÜST PROFİL VE OTURUM KAPATMA ---
        ust_bar_duzeni = QHBoxLayout()

        self.btn_profil = QPushButton("👤 Giriş Yapılmadı")
        self.btn_profil.setStyleSheet(
            "background-color: transparent; color: #2C3E50; font-weight: bold; font-size: 16px; text-align: left;")
        self.btn_profil.setToolTip("Profili Düzenle")
        self.btn_profil.clicked.connect(self.profil_ac)

        btn_cikis = QPushButton("Güvenli Çıkış")
        btn_cikis.setStyleSheet(
            "background-color: #C0392B; color: white; font-weight: bold; padding: 5px 15px; border-radius: 4px;")
        btn_cikis.clicked.connect(self.cikis_yap)

        ust_bar_duzeni.addWidget(self.btn_profil)
        ust_bar_duzeni.addStretch()
        ust_bar_duzeni.addWidget(btn_cikis)

        ana_duzen.addLayout(ust_bar_duzeni)
        # ----------------------------------------

        ana_duzen.setContentsMargins(30, 30, 30, 30)

        # ---------------- 1. ÜST BİLGİ ALANI (HEADER) ----------------
        header_duzen = QHBoxLayout()

        self.logo_etiketi = QLabel()
        self.logo_etiketi.setPixmap(QPixmap(kaynak_yolu("arayuz/varliklar/aemot_logo_gorseli.png")))
        self.logo_etiketi.setScaledContents(True)
        self.logo_etiketi.setFixedSize(180, 50)
        self.logo_etiketi.setStyleSheet("background: transparent; border: none;")

        self.saat_etiketi = QLabel("SAAT YÜKLENİYOR...")
        self.saat_etiketi.setStyleSheet(
            "font-size: 16px; font-weight: bold; color: #2C3E50; background: transparent; border: none;")
        self.saat_etiketi.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        header_duzen.addWidget(self.logo_etiketi)
        header_duzen.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        header_duzen.addWidget(self.saat_etiketi)

        ana_duzen.addLayout(header_duzen)
        ana_duzen.addSpacing(40)  # Logo ile orta kart arasına biraz boşluk

        # ---------------- 2. ORTA MENÜ KARTI ----------------
        kart_frame = QFrame()
        kart_frame.setObjectName("KartFrame")  # CSS'te özel yakalamak için ID verdik

        # Kartın ne kadar genişleyeceğini sınırlayalım ki ekrana yayılıp çirkin durmasın
        kart_frame.setMinimumWidth(500)
        kart_frame.setMaximumWidth(600)

        kart_duzen = QVBoxLayout(kart_frame)
        kart_duzen.setContentsMargins(40, 40, 40, 40)
        kart_duzen.setSpacing(20)

        baslik = QLabel("AEMOT ÜRETİM VE PLANLAMA SİSTEMİ")
        baslik.setAlignment(Qt.AlignmentFlag.AlignCenter)
        baslik.setStyleSheet(
            "font-size: 22px; font-weight: 900; color: #2C3E50; margin-bottom: 20px; background: transparent; border: none;")
        kart_duzen.addWidget(baslik)

        # Butonlar
        btn_id_motor = QPushButton("⚙️  1. Motor ID ve Açıklama Oluşturucu")
        btn_id_motor.setObjectName("btn1")
        btn_id_motor.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_id_motor.clicked.connect(lambda: ana_pencere.sayfa_degistir(1))

        btn_urun_agaci = QPushButton("📦  2. Ham Madde Ürün Ağacı (Yakında)")
        btn_urun_agaci.setObjectName("btn2")
        btn_urun_agaci.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_urun_agaci.clicked.connect(lambda: ana_pencere.sayfa_degistir(2))

        btn_ayarlar = QPushButton("🛠️  3. Ayarlar ve Dosya Yolları")
        btn_ayarlar.setObjectName("btn3")
        btn_ayarlar.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_ayarlar.clicked.connect(lambda: ana_pencere.sayfa_degistir(3))

        # (Ana Menüdeki diğer butonları eklediğin yerin altına koy)
        self.btn_yardim = QPushButton("❓ Nasıl Kullanılır?")
        self.btn_yardim.setStyleSheet(
            "background-color: #F39C12; color: white; font-weight: bold; padding: 12px; border-radius: 6px; font-size: 14px;")
        self.btn_yardim.clicked.connect(self.kullanim_kilavuzu_goster)

        # self.buton_layout.addWidget(self.btn_yardim) # Ana menündeki layout'un adı neyse ona ekle.

        kart_duzen.addWidget(btn_id_motor)
        kart_duzen.addWidget(btn_urun_agaci)
        kart_duzen.addWidget(btn_ayarlar)
        kart_duzen.addWidget(self.btn_yardim)

        # Kartı ana düzene tam ortaya (AlignHCenter) hizalayarak ekle
        ana_duzen.addWidget(kart_frame, alignment=Qt.AlignmentFlag.AlignHCenter)

        # Karttan sonra boşluk bırakarak kartı yukarı doğru it
        ana_duzen.addStretch()

        # ---------------- 3. ALT BİLGİ (FOOTER) ----------------
        footer = QLabel("Muhammet Ali Aslanhan | AEMOT Staj Projesi")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet(
            "font-size: 12px; font-weight: bold; color: #7F8C8D; background: transparent; border: none;")
        ana_duzen.addWidget(footer)

        # ---------------- SAAT ZAMANLAYICISI ----------------
        self.zamanlayici = QTimer(self)
        self.zamanlayici.timeout.connect(self.tarih_mevcut_gosterim)
        self.zamanlayici.start(1000)
        self.tarih_mevcut_gosterim()  # Ekran açılır açılmaz saati göstersin diye 1 kere manuel çağırıyoruz

    def tarih_mevcut_gosterim(self):
        suan = QDateTime.currentDateTime()
        metin = suan.toString("dd.MM.yyyy - HH:mm:ss")
        self.saat_etiketi.setText(metin)

    def kullanim_kilavuzu_goster(self):
        from PyQt6.QtWidgets import QMessageBox
        kilavuz_metni = (
            "📌 VERTEX - AEMOT YÖNETİM PANELİ KULLANIM KILAVUZU\n\n"
            "1️⃣ VERİTABANI BAĞLANTISI:\n"
            "Öncelikle 'Ayarlar' sekmesine giderek kodların kaydedileceği Ortak Excel dosyasını seçin. Seçtiğiniz yol kalıcı olarak sisteme kaydedilir.\n\n"
            "2️⃣ YENİ BİR KOD ÜRETMEK:\n"
            "Motor, Flanş veya Kapak arayüzüne girin. Özellikleri seçin. Yanlış/Eksik veya standart dışı bir bilgi girerseniz sistem sizi uyaracaktır.\n\n"
            "3️⃣ YENİ BİR STANDART EKLEME (ÖNEMLİ!):\n"
            "AEMOT yeni bir Gövde/Kutup kombinasyonu çıkardıysa, kimseye ihtiyaç duymadan 'Ayarlar' sekmesindeki butonları kullanarak sisteme anında öğretebilirsiniz.\n\n"
            "4️⃣ EXCEL'E AKTARMA VE GÜVENLİK:\n"
            "Tablodaki verileri 'Tümünü Excel'e Aktar' butonuyla gönderdiğinizde, sistem arka planda Excel'in o anki yedeğini 'AEMOT_AYARLAR/Yedekler' klasörüne otomatik olarak kaydeder. Bir hata olursa oradan kurtarabilirsiniz.\n\n"
            "🚨 KRİTİK UYARI: Excel dosyanızdaki Sütun Başlıklarının adlarını (KOD, AÇIKLAMA vs.) KESİNLİKLE DEĞİŞTİRMEYİN! (Programın arama motoru bu başlıklara göre çalışır.)"
        )
        QMessageBox.information(self, "Kullanım Kılavuzu", kilavuz_metni)

    def profil_guncelle(self):
        """Kullanıcı giriş yaptığında bu fonksiyon tetiklenip ismi yazar."""
        self.btn_profil.setText(f"👤 {oturum.tam_isim()}")

    def profil_ac(self):
        """Profil butonuna tıklanınca şifre değiştirme ekranı açılır."""
        dialog = ProfilEkrani(self)
        dialog.exec()

    def cikis_yap(self):
        cevap = QMessageBox.question(self, "Güvenli Çıkış", "Oturumu kapatmak istediğinize emin misiniz?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if cevap == QMessageBox.StandardButton.Yes:
            oturum.aktif_kimlik = ""
            oturum.aktif_ad = ""
            oturum.aktif_soyad = ""

            # Ana pencereye "Kullanıcı çıkış yaptı" bayrağını dikiyoruz ve dev pencereyi kapatıyoruz.
            # Pencere kapanınca main.py içindeki döngü bunu algılayıp sistemi başa (Giriş Ekranına) saracak.
            self.ana_pencere.cikis_yapildi = True
            self.ana_pencere.close()