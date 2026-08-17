import json
import os
import sys
import shutil
import datetime

# --- 1. ORTAK AĞ KLASÖRÜNE KAYIT (TÜM PC'LER SENKRONİZE OLSUN DİYE) ---
# Artık program neredeyse (.exe veya ana dosya), veritabanı klasörü orasıdır.
ANA_DIZIN = os.path.abspath(".")
KALICI_KLASOR = os.path.join(ANA_DIZIN, "VERITABANI_VERTEX")

if not os.path.exists(KALICI_KLASOR):
    os.makedirs(KALICI_KLASOR)

if not os.path.exists(KALICI_KLASOR):
    os.makedirs(KALICI_KLASOR)

# --- 2. DOSYALARI BELGELER/VERTEX İÇİNE BAĞLAMA ---
birlestir = os.path.join(KALICI_KLASOR, 'katalog.json')
flans_birlestir = os.path.join(KALICI_KLASOR, 'standart_flans.json')
cap_birlestir = os.path.join(KALICI_KLASOR, 'cap.json')
ayarlar_dosyasi = os.path.join(KALICI_KLASOR, 'ayarlar.json')


# --- 3. İLK KURULUMDA ŞABLONLARI KURTARMA ---
def ilk_kurulum_kopyalamasi():
    if getattr(sys, 'frozen', False):
        GIZLI_DIZIN = sys._MEIPASS
    else:
        GIZLI_DIZIN = os.path.abspath(".")

    gomulu_katalog = os.path.join(GIZLI_DIZIN, 'arayuz', 'mantik', 'katalog_veri_json', 'katalog.json')
    gomulu_flans = os.path.join(GIZLI_DIZIN, 'arayuz', 'mantik', 'katalog_veri_json', 'standart_flans.json')
    gomulu_cap = os.path.join(GIZLI_DIZIN, 'arayuz', 'mantik', 'katalog_veri_json', 'cap.json')

    try:
        if not os.path.exists(birlestir) and os.path.exists(gomulu_katalog):
            shutil.copy(gomulu_katalog, birlestir)
        if not os.path.exists(flans_birlestir) and os.path.exists(gomulu_flans):
            shutil.copy(gomulu_flans, flans_birlestir)
        if not os.path.exists(cap_birlestir) and os.path.exists(gomulu_cap):
            shutil.copy(gomulu_cap, cap_birlestir)
    except Exception:
        pass


ilk_kurulum_kopyalamasi()


# ==========================================================
# İŞLEM FONKSİYONLARI (Hepsi Belgeler/VERTEX'ten okur/yazar)
# ==========================================================

def cap_oku():
    try:
        with open(cap_birlestir, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def cap_yaz(veri):
    with open(cap_birlestir, 'w', encoding='utf-8') as f: json.dump(veri, f, ensure_ascii=False, indent=4)


def katalog_oku():
    try:
        with open(birlestir, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def katalog_yaz(veri):
    with open(birlestir, 'w', encoding='utf-8') as f: json.dump(veri, f, ensure_ascii=False, indent=4)


def standart_flans_oku():
    try:
        with open(flans_birlestir, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def standart_flans_yaz(veri):
    with open(flans_birlestir, 'w', encoding='utf-8') as f: json.dump(veri, f, ensure_ascii=False, indent=4)


# --- 4. EXCEL YOLU ---
def ayar_excel_yolunu_getir():
    varsayilan_yol = r'C:\Katalog_Secilmedi.xlsx'
    if not os.path.exists(ayarlar_dosyasi):
        return varsayilan_yol
    try:
        with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
            ayarlar = json.load(f)
            return ayarlar.get("excel_dosya_yolu", varsayilan_yol)
    except Exception:
        return varsayilan_yol


def ayar_excel_yolunu_kaydet(yeni_yol):
    try:
        ayarlar = {}
        if os.path.exists(ayarlar_dosyasi):
            with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
                ayarlar = json.load(f)

        ayarlar["excel_dosya_yolu"] = yeni_yol

        with open(ayarlar_dosyasi, "w", encoding="utf-8") as f:
            json.dump(ayarlar, f, ensure_ascii=False, indent=4)
        return True
    except Exception:
        return False


def urun_agaci_excel_yolunu_getir():
    varsayilan_yol = r'C:\Urun_Agaci_Kayit_Bulunamadi.xlsx'
    if not os.path.exists(ayarlar_dosyasi):
        return varsayilan_yol
    try:
        with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
            ayarlar = json.load(f)
            return ayarlar.get("urun_agaci_excel_yolu", varsayilan_yol)
    except Exception:
        return varsayilan_yol


def urun_agaci_excel_yolunu_kaydet(yeni_yol):
    try:
        ayarlar = {}
        if os.path.exists(ayarlar_dosyasi):
            with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
                ayarlar = json.load(f)

        ayarlar["urun_agaci_excel_yolu"] = yeni_yol

        with open(ayarlar_dosyasi, "w", encoding="utf-8") as f:
            json.dump(ayarlar, f, ensure_ascii=False, indent=4)
        return True
    except Exception:
        return False


# --- 5. OTOMATİK YEDEKLEME (SENİN İSTEDİĞİN FORMATTA) ---
def excel_yedekle(excel_yolu):
    if not os.path.exists(excel_yolu):
        return

    yedek_klasor = os.path.join(KALICI_KLASOR, "Yedekler")
    if not os.path.exists(yedek_klasor):
        os.makedirs(yedek_klasor)

    # Son 30 yedeği tutarak diskin dolmasını engelliyoruz
    mevcut_yedekler = sorted([f for f in os.listdir(yedek_klasor) if f.endswith(".xlsx")])
    if len(mevcut_yedekler) > 30:
        try:
            os.remove(os.path.join(yedek_klasor, mevcut_yedekler[0]))
        except:
            pass

    # SENİN İSTEDİĞİN FORMAT: tarih_dosyaninadi.xlsx
    zaman = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
    dosya_adi = os.path.basename(excel_yolu)
    yedek_adi = f"{zaman}_{dosya_adi}"

    try:
        shutil.copy(excel_yolu, os.path.join(yedek_klasor, yedek_adi))
    except Exception:
        pass

    # --- 6. HATA LOGLAMA ---


def hata_logla(hata_metni):
    log_dosyasi = os.path.join(KALICI_KLASOR, "hata_kayit.txt")
    zaman = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    try:
        with open(log_dosyasi, "a", encoding="utf-8") as f:
            f.write(f"[{zaman}] HATA:\n{hata_metni}\n{'-' * 60}\n")
    except:
        pass

def kaynak_yolu(goreceli_yol):
    """ PyInstaller ile oluşturulan Exe içindeki varlıkların yolunu bulur """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, goreceli_yol)


def log_aktif_mi():
    """ayarlar.json'dan log_aktif bayrağını okur. Varsayılan: True"""
    try:
        if os.path.exists(ayarlar_dosyasi):
            with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
                ayarlar = json.load(f)
                return ayarlar.get("log_aktif", True)
    except Exception:
        pass
    return True


def log_aktif_ayarla(deger: bool):
    """Log sistemini aktif veya deaktif eder. ayarlar.json'a kaydeder."""
    try:
        ayarlar = {}
        if os.path.exists(ayarlar_dosyasi):
            with open(ayarlar_dosyasi, "r", encoding="utf-8") as f:
                ayarlar = json.load(f)
        ayarlar["log_aktif"] = deger
        with open(ayarlar_dosyasi, "w", encoding="utf-8") as f:
            json.dump(ayarlar, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


def log_kayitlarini_sil():
    """Tüm sistem log kayıtlarını siler. Sadece as yetkili kullanabilir."""
    try:
        log_dosyasi = os.path.join(KALICI_KLASOR, "sistem_gecmisi.json")
        with open(log_dosyasi, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=4)
        return True
    except Exception:
        return False


def sistem_logu_ekle(mesaj):
    """Log toggle kapalıysa kayıt yapmaz. Açıksa kayıt ekler."""
    # Toggle kontrolü: kapalıysa hiçbir şey yapma
    if not log_aktif_mi():
        return
    try:
        log_dosyasi = os.path.join(KALICI_KLASOR, "sistem_gecmisi.json")
        zaman = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        log_kaydi = f"[{zaman}] {mesaj}"

        loglar = []
        if os.path.exists(log_dosyasi):
            with open(log_dosyasi, "r", encoding="utf-8") as f:
                loglar = json.load(f)

        loglar.insert(0, log_kaydi)  # En son yapılan işlem en üste gelsin
        if len(loglar) > 1000: loglar = loglar[:1000]  # Dosya çok şişmesin diye son 1000 kayıt tutulur

        with open(log_dosyasi, "w", encoding="utf-8") as f:
            json.dump(loglar, f, ensure_ascii=False, indent=4)
    except:
        pass