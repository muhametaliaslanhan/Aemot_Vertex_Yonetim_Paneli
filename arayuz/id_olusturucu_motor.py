from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QPushButton, QStackedWidget,
    QListWidget, QFrame, QLabel, QSizePolicy, QListWidgetItem
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QCursor

from arayuz.parcalar.rotorlu_mil import RotorluMilArayuzu
from arayuz.parcalar.ana_motor import MotorArayuzu
from arayuz.parcalar.rotor_saci import RotorSaciArayuzu
from arayuz.parcalar.rotor_paket import RotorPaketArayuzu
from arayuz.parcalar.mil import MilArayuzu
from arayuz.parcalar.statorlu_govde import StatorluGovdeArayuzu
from arayuz.parcalar.bobinli_stator import BobinliStatorArayuzu
from arayuz.parcalar.sarilmis_bobin import SarilmisBobinArayuzu
from arayuz.parcalar.stator_paket import StatorPaketArayuzu
from arayuz.parcalar.stator_saci import  StatorSaciArayuzu
from arayuz.parcalar.govde import GovdeArayuzu
from arayuz.parcalar.govde_ayak import GovdeAyakArayuzu
from arayuz.parcalar.govde_ayak_takim import GovdeAyakTakimArayuzu
from arayuz.parcalar.stator_rotor_pul import StatorRotorPulArayuzu
from arayuz.parcalar.klemens_kutu_kapagi import Klemens_Kutusu_kapagiArayuzu
from arayuz.parcalar.flans import FlansArayuzu
from arayuz.parcalar.kapak import KapakArayuzu
from arayuz.parcalar.rulman_yag_kapagi import RulmanYagKapagiArayuzu
from arayuz.parcalar.fan_muhafaza import FanMuhafazaArayuzu
from arayuz.parcalar.kanopi import KanopiArayuzu

class UretimAnaPaneli(QWidget):
    def __init__(self, ana_pencere):
        super().__init__()
        self.ana_pencere = ana_pencere

        # Genel Stil (Modern ve Sade)
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', Arial; }
            QFrame#YanMenu { background-color: #2C3E50; border-right: 2px solid #34495E; }
            QListWidget { background-color: transparent; border: none; outline: none; }
            QListWidget::item { color: #ECF0F1; padding: 12px; font-size: 14px; font-weight: bold; border-bottom: 1px solid #34495E; }
            QListWidget::item:selected { background-color: #E74C3C; border-left: 5px solid #C0392B; }
            QListWidget::item:hover { background-color: #34495E; }
            QPushButton#BtnToggle { background-color: #E74C3C; color: white; font-weight: bold; font-size: 18px; border: none; border-radius: 0px; }
            QPushButton#BtnToggle:hover { background-color: #C0392B; }
        """)

        ana_duzen = QHBoxLayout(self)
        ana_duzen.setContentsMargins(0, 0, 0, 0)
        ana_duzen.setSpacing(0)

        # ---------------- 1. SOL YAN MENÜ (SIDEBAR) ----------------
        self.yan_menu_frame = QFrame()
        self.yan_menu_frame.setObjectName("YanMenu")
        self.yan_menu_frame.setFixedWidth(250)  # Menünün genişliği

        yan_menu_duzen = QVBoxLayout(self.yan_menu_frame)
        yan_menu_duzen.setContentsMargins(0, 20, 0, 0)

        menu_baslik = QLabel("ÜRETİM PARÇALARI")
        menu_baslik.setStyleSheet(
            "color: #95A5A6; font-size: 12px; font-weight: bold; padding-left: 15px; margin-bottom: 10px;")
        yan_menu_duzen.addWidget(menu_baslik)

        # Liste ve Türkçe İsimler
        self.parca_listesi = QListWidget()
        self.parca_listesi.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        # Kullanıcının göreceği Türkçe ve düzgün isimler
        liste_elemanlari = [
            "1. Ana Motor", "2. Rotorlu Mil", "3. Mil", "4. Rotor Paket", "5. Rotor Sacı",
            "6. Statorlu Gövde", "7. Bobinli Stator", "8. Sarılmış Bobin", "9. Stator Paket",
            "10. Stator Sacı", "11. Gövde", "12. Gövde Ayak", "13. Gövde Ayak Takım",
            "14. Stator-Rotor Pul", "15. Klemens Kutu & Kapak", "16. Flanş", "17. Kapak",
            "18. Rulman Yağ Kapağı", "19. Fan Muhafazası", "20. Kanopi", "21. Leroy Somer",
            "22. Agraf Sacı"
        ]

        for isim in liste_elemanlari:
            item = QListWidgetItem(isim)
            self.parca_listesi.addItem(item)

        yan_menu_duzen.addWidget(self.parca_listesi)

        # ---------------- 2. AÇ/KAPA (TOGGLE) BUTONU ----------------
        # Senin o ince hat üstündeki ok dediğin buton bu
        self.btn_toggle = QPushButton("◀")
        self.btn_toggle.setObjectName("BtnToggle")
        self.btn_toggle.setFixedSize(25, 100)  # İnce uzun bir buton
        self.btn_toggle.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_toggle.setToolTip("Menüyü Gizle/Göster")
        self.btn_toggle.clicked.connect(self.menuyu_ac_kapa)

        # ---------------- 3. SAĞ İÇERİK ALANI (STACKED WIDGET) ----------------
        self.sayfalar = QStackedWidget()

        # Sayfa 0: Ana Motor (İndeks 0)
        self.sayfa_ana_motor = MotorArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_ana_motor)

        # Sayfa 1: Rotorlu Mil (İndeks 1)  <-- BUNU YENİ BAĞLADIK
        self.sayfa_rotorlu_mil = RotorluMilArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_rotorlu_mil)

        #mil bağlandı
        self.sayfa_mil = MilArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_mil)

        self.sayfa_rotor_paket = RotorPaketArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_rotor_paket)

        self.sayfa_rotor_saci = RotorSaciArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_rotor_saci)

        self.sayfa_statorlu_govde = StatorluGovdeArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_statorlu_govde)

        self.sayfa_bobinli_stator = BobinliStatorArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_bobinli_stator)

        self.sayfa_sarilmis_bobin = SarilmisBobinArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_sarilmis_bobin)

        self.sayfa_stator_paket = StatorPaketArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_stator_paket)

        self.sayfa_stator_saci = StatorSaciArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_stator_saci)

        self.sayfa_govde = GovdeArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_govde)

        self.sayfa_govde_ayak = GovdeAyakArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_govde_ayak)

        self.sayfa_govde_ayak_takim = GovdeAyakTakimArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_govde_ayak_takim)

        self.sayfa_stator_rotor_pul = StatorRotorPulArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_stator_rotor_pul)

        self.sayfa_klemens_kutu_kapagi = Klemens_Kutusu_kapagiArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_klemens_kutu_kapagi)

        self.sayfa_flans = FlansArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_flans)

        self.sayfa_kapak = KapakArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_kapak)

        self.sayfa_rulman_yag_kapagi = RulmanYagKapagiArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_rulman_yag_kapagi)

        self.sayfa_fan_muhafaza = FanMuhafazaArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_fan_muhafaza)

        self.sayfa_kanopi = KanopiArayuzu(self.ana_pencere)
        self.sayfalar.addWidget(self.sayfa_kanopi)


        # Diğer kalan 20 parça için şimdilik boş "Yapım Aşamasında" sayfaları (Artık 2. indeksten başlıyor)
        for i in range(20, len(liste_elemanlari)):
            bos_sayfa = QWidget()
            bos_duzen = QVBoxLayout(bos_sayfa)
            mesaj = QLabel(f"{liste_elemanlari[i]}\n(Bu modül yakında eklenecektir)")
            mesaj.setAlignment(Qt.AlignmentFlag.AlignCenter)
            mesaj.setStyleSheet("font-size: 24px; font-weight: bold; color: #7F8C8D;")
            bos_duzen.addWidget(mesaj)
            self.sayfalar.addWidget(bos_sayfa)

        # Listeden seçim yapınca sağdaki sayfayı değiştir
        self.parca_listesi.currentRowChanged.connect(self.sayfa_degistir)
        self.parca_listesi.setCurrentRow(0)  # Başlangıçta Ana Motor seçili olsun

        # ---------------- EKRANA YERLEŞTİRME ----------------
        ana_duzen.addWidget(self.yan_menu_frame)

        # Butonu ve Stacked Widget'ı bir araya getiren alt düzen
        sag_duzen = QHBoxLayout()
        sag_duzen.setContentsMargins(0, 0, 0, 0)
        sag_duzen.setSpacing(0)

        sag_duzen.addWidget(self.btn_toggle)
        sag_duzen.addWidget(self.sayfalar)

        ana_duzen.addLayout(sag_duzen)

    def menuyu_ac_kapa(self):
        """Menü görünürse gizler, gizliyse gösterir ve ok yönünü değiştirir."""
        if self.yan_menu_frame.isVisible():
            self.yan_menu_frame.setVisible(False)
            self.btn_toggle.setText("▶")
        else:
            self.yan_menu_frame.setVisible(True)
            self.btn_toggle.setText("◀")

    def sayfa_degistir(self, indeks):
        """Sol menüden tıklanan sıraya göre sağdaki sayfayı açar."""
        self.sayfalar.setCurrentIndex(indeks)