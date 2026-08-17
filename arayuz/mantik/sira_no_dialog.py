# arayuz/mantik/sira_no_dialog.py
"""
Ortak Sıra No Seçim Paneli Modülü
===================================
parcalar klasöründeki tüm dosyalar için tek bir kaynak dialog.
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QInputDialog, QLineEdit
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont


def _sayfa_bul(xls, sayfa_adi):
    """Verilen sayfa_adi'na önce birebir, sonra kelime-içerme ile en doğru Excel sayfasını bulur."""
    if not sayfa_adi:
        return None

    hedef_norm = str(sayfa_adi).strip().upper()

    # 1) Birebir eşleşme
    for sayfa in xls.sheet_names:
        if sayfa.strip().upper() == hedef_norm:
            return sayfa

    # 2) Kelime bazlı tam içerme (Örn: "STATOR PAKET" -> sayfa adında hem STATOR hem PAKET geçmeli)
    kelimeler = [k for k in hedef_norm.replace("-", " ").split() if k]
    if kelimeler:
        for sayfa in xls.sheet_names:
            sayfa_norm = sayfa.strip().upper()
            if all(k in sayfa_norm for k in kelimeler):
                return sayfa

    return None


def _normalize_sozluk(parent_widget, dna_sonuc, hedef_metin, sayfa_adi=None, id_kodu=None):
    """
    Sadece INT (Örn: 182) geldiyse, arka planda Excel'de DOĞRU sayfada, DOĞRU
    sütunlarda (ID ve 'AÇIKLAMA SIRA NO') KESİN eşleşme arayarak ID, adet ve
    Ana Açıklamayı bulur.

    Önceki sürüm tüm sayfaları/hücreleri metin olarak körlemesine tarıyor ve
    "-0182" alt dizisini "-01820" gibi metinlerde de eşleştiriyordu; bu da
    yanlış sayfadan/parça tipinden kayıt gelmesine ve yanlış adet/kod/açıklama
    gösterilmesine yol açıyordu. Bu sürüm, çağıran dosyanın zaten bildiği
    sayfa_adi ve id_kodu bilgisini kullanarak sadece doğru sayfada, sadece
    'AÇIKLAMA SIRA NO' sütunundaki SAYISAL değeri sira_int ile birebir
    karşılaştırır (alt dize araması değil).
    """
    if dna_sonuc is None:
        return {}

    if isinstance(dna_sonuc, dict):
        return dna_sonuc

    if not (isinstance(dna_sonuc, int) or str(dna_sonuc).isdigit()):
        return {}

    sira_int = int(dna_sonuc)
    ornek_id = "GEÇMİŞ BULUNAMADI"
    ornek_aciklama = hedef_metin
    adet = 0

    try:
        from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir
        import pandas as pd

        excel_yolu = ayar_excel_yolunu_getir()
        xls = pd.ExcelFile(excel_yolu)

        hedef_sayfa = _sayfa_bul(xls, sayfa_adi)
        # sayfa_adi verilmemişse veya hiç bulunamazsa (beklenmeyen durum),
        # eski davranışa geri dönmemek için tüm sayfaları tarıyoruz ama YİNE
        # DE sadece sayısal sütun eşleşmesiyle - metin taraması yapmıyoruz.
        sayfalar = [hedef_sayfa] if hedef_sayfa else list(xls.sheet_names)

        id_kodu_norm = str(id_kodu).strip().upper() if id_kodu else None

        for sayfa in sayfalar:
            try:
                df = pd.read_excel(excel_yolu, sheet_name=sayfa, header=0, dtype=str).fillna("")
            except Exception:
                continue
            df.columns = [str(c).strip().upper() for c in df.columns]

            # ID / KOD sütunu (Örn: "#KOD")
            id_kolonu = None
            for col in df.columns:
                if "KOD" in col and "AÇIKLAMA" not in col:
                    id_kolonu = col
                    break

            # Sıra No sütunu - önce tam adıyla, yoksa "SIRA" ve "NO" geçen (RESİM hariç) sütun
            sira_kolonu = None
            if 'AÇIKLAMA SIRA NO' in df.columns:
                sira_kolonu = 'AÇIKLAMA SIRA NO'
            else:
                for col in df.columns:
                    if "SIRA" in col and "NO" in col and "RESİM" not in col:
                        sira_kolonu = col
                        break

            if not sira_kolonu:
                continue

            aciklama_kolonu = 'AÇIKLAMA' if 'AÇIKLAMA' in df.columns else None
            aciklama2_kolonu = 'AÇIKLAMA2' if 'AÇIKLAMA2' in df.columns else (
                'EK AÇIKLAMA' if 'EK AÇIKLAMA' in df.columns else None
            )

            for _, row in df.iterrows():
                ham_sira = str(row.get(sira_kolonu, "")).replace("'", "").strip()
                if not ham_sira:
                    continue
                try:
                    satir_sira_int = int(float(ham_sira))
                except (ValueError, TypeError):
                    continue

                # KESİN SAYISAL EŞLEŞME - alt dize değil, birebir sayı karşılaştırması
                if satir_sira_int != sira_int:
                    continue

                satir_id = str(row.get(id_kolonu, "")).strip().upper() if id_kolonu else ""

                # id_kodu verilmişse (Örn: "SP", "FM") ve ID'de bu kod yoksa,
                # bu satır farklı bir parça tipine ait demektir - atla.
                if id_kodu_norm and satir_id and f"-{id_kodu_norm}-" not in satir_id and f"-{id_kodu_norm}" not in satir_id:
                    continue

                adet += 1
                if ornek_id == "GEÇMİŞ BULUNAMADI":
                    if satir_id:
                        ornek_id = satir_id
                    parcalar_aciklama = []
                    if aciklama_kolonu:
                        deger = str(row.get(aciklama_kolonu, "")).strip()
                        if deger:
                            parcalar_aciklama.append(deger)
                    if aciklama2_kolonu:
                        deger2 = str(row.get(aciklama2_kolonu, "")).strip()
                        if deger2:
                            parcalar_aciklama.append(deger2)
                    if parcalar_aciklama:
                        ornek_aciklama = " ".join(parcalar_aciklama)

            # Doğru sayfa zaten belliydi ve bu sayfayı taradık -> döngüyü bitir
            if hedef_sayfa:
                break
            # sayfa_adi verilmemişti ama bir sayfada eşleşme bulundu -> yeter
            if adet > 0:
                break

    except Exception as e:
        print("Sıra No Dialog Excel Arama Hatası:", e)

    return {
        sira_int: {
            'adet': max(1, adet),
            'tam_aciklama': ornek_aciklama,
            'id': ornek_id
        }
    }


class SiraNoSecimPaneli(QDialog):
    def __init__(self, parent, bulunanlar_sozlugu: dict, aciklama_metni: str):
        super().__init__(parent)
        self.setWindowTitle("Sıra Numarası Çakışması / Seçimi")
        self.setMinimumSize(870, 420)
        self.setStyleSheet("""
            QDialog { background-color: #F8F9FA; }
            QLabel { color: #2C3E50; }
            QTableWidget {
                font-size: 13px;
                border: 1px solid #BDC3C7;
                border-radius: 5px;
                gridline-color: #ECF0F1;
            }
            QTableWidget::item { padding: 6px; }
            QTableWidget::item:selected { background-color: #3498DB; color: white; }
            QHeaderView::section {
                background-color: #2C3E50;
                color: white;
                font-weight: bold;
                padding: 6px;
                border: none;
            }
            QPushButton {
                font-weight: bold;
                padding: 10px 16px;
                border-radius: 5px;
                border: none;
                font-size: 13px;
            }
        """)

        self._secilen_no = None
        ana_duzen = QVBoxLayout(self)
        ana_duzen.setSpacing(10)
        ana_duzen.setContentsMargins(14, 14, 14, 14)

        if bulunanlar_sozlugu:
            info_metni = (
                f"Bu ürün için geçmişte <b>{len(bulunanlar_sozlugu)}</b> farklı sıra numarası kullanılmış.\n"
                f"Açıklama: <i>{aciklama_metni[:120]}</i>\n\n"
                f"Lütfen tablodan bir numara seçin, yeni numara üretin veya elle girin."
            )
        else:
            info_metni = (
                "Bu ürün için geçmişte eşleşen bir kayıt <b>bulunamadı</b>.\n"
                "Yeni sıra numarası üretebilir veya elle girebilirsiniz."
            )

        lbl_info = QLabel(info_metni)
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet("font-size: 13px; font-weight: bold; color: #C0392B; margin-bottom: 4px;")
        lbl_info.setTextFormat(Qt.TextFormat.RichText)
        ana_duzen.addWidget(lbl_info)

        self.tablo = QTableWidget(len(bulunanlar_sozlugu), 4)
        self.tablo.setHorizontalHeaderLabels(["Sıra No", "Kullanım Adedi", "Örnek Parça Kodu", "Ana Açıklama (Örnek)"])
        self.tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tablo.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.tablo.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tablo.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.tablo.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tablo.setAlternatingRowColors(True)

        siralananlar = sorted(bulunanlar_sozlugu.keys())
        for satir_idx, no in enumerate(siralananlar):
            detay = bulunanlar_sozlugu[no]

            item_no = QTableWidgetItem(str(no).zfill(4))
            item_no.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_no.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))

            item_adet = QTableWidgetItem(f"{detay.get('adet', 1)} Adet")
            item_adet.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_adet.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            item_adet.setForeground(Qt.GlobalColor.darkGreen)

            item_id = QTableWidgetItem(str(detay.get('id', '')))
            item_id.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            item_aciklama = QTableWidgetItem(str(detay.get('tam_aciklama', '')))

            self.tablo.setItem(satir_idx, 0, item_no)
            self.tablo.setItem(satir_idx, 1, item_adet)
            self.tablo.setItem(satir_idx, 2, item_id)
            self.tablo.setItem(satir_idx, 3, item_aciklama)

        if bulunanlar_sozlugu:
            self.tablo.selectRow(0)

        ana_duzen.addWidget(self.tablo, stretch=1)

        buton_duzen = QHBoxLayout()

        btn_elle_gir = QPushButton("🔒 Elle Sıra No Gir")
        btn_elle_gir.setStyleSheet("background-color: #8E44AD; color: white;")
        btn_elle_gir.clicked.connect(self._on_elle_gir)

        btn_sec = QPushButton("Seçili Numarayı Uygula")
        btn_sec.setStyleSheet("background-color: #27AE60; color: white;")
        btn_sec.clicked.connect(self._on_sec)
        btn_sec.setEnabled(bool(bulunanlar_sozlugu))

        btn_yeni = QPushButton("Yeni Sıra No Üret")
        btn_yeni.setStyleSheet("background-color: #3498DB; color: white;")
        btn_yeni.clicked.connect(self._on_yeni_uret)

        btn_iptal = QPushButton("İptal")
        btn_iptal.setStyleSheet("background-color: #E74C3C; color: white;")
        btn_iptal.clicked.connect(self.reject)

        buton_duzen.addWidget(btn_elle_gir)
        buton_duzen.addStretch()
        buton_duzen.addWidget(btn_sec)
        buton_duzen.addWidget(btn_yeni)
        buton_duzen.addWidget(btn_iptal)
        ana_duzen.addLayout(buton_duzen)

    def _on_sec(self):
        secili_satirlar = self.tablo.selectionModel().selectedRows()
        if secili_satirlar:
            secilen_no_str = self.tablo.item(secili_satirlar[0].row(), 0).text()
            self._secilen_no = int(secilen_no_str)
            self.accept()
        else:
            QMessageBox.warning(self, "Uyarı", "Lütfen tablodan bir numara seçin!")

    def _on_yeni_uret(self):
        self._secilen_no = -1
        self.accept()

    def _on_elle_gir(self):
        sifre, ok = QInputDialog.getText(
            self, "Yetkili Girişi",
            "Lütfen yetkili şifresini girin:",
            QLineEdit.EchoMode.Password
        )
        if ok and sifre == "aemot123":
            no, ok2 = QInputDialog.getInt(
                self, "Elle Sıra No Gir",
                "Atamak istediğiniz sıra numarasını girin:",
                min=1, max=99999
            )
            if ok2:
                self._secilen_no = no
                self.accept()
        elif ok:
            QMessageBox.warning(self, "Yetkisiz Erişim", "Hatalı şifre girdiniz!")

    def exec_ve_dondur(self):
        sonuc = self.exec()
        if sonuc == QDialog.DialogCode.Accepted:
            return self._secilen_no
        return None


def sira_no_secim_yap(parent_widget, dna_sonuc, aciklama_metni: str, sayfa_adi: str = None, id_kodu: str = None):
    """
    sayfa_adi: İlgili Excel sayfasının adı (Örn: "STATOR PAKET", "FAN MUHAFAZA").
               Verilmesi şiddetle önerilir - aksi halde sayfa tahmini yapılmaya
               çalışılır ve yanlış sayfadan kayıt gelme riski oluşur.
    id_kodu:   İlgili parçanın ID kodu (Örn: "SP", "FM", "GV") - aynı sayfada
               farklı kod öneklerine sahip kayıtları elemek için ek güvenlik katmanıdır.
    """
    bulunanlar = _normalize_sozluk(parent_widget, dna_sonuc, aciklama_metni, sayfa_adi=sayfa_adi, id_kodu=id_kodu)
    panel = SiraNoSecimPaneli(parent_widget, bulunanlar, aciklama_metni)
    return panel.exec_ve_dondur()
