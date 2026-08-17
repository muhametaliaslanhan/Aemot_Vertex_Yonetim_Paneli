from arayuz.mantik.json_islemleri import katalog_oku, katalog_yaz, standart_flans_oku,standart_flans_yaz

class motor_ozellikleri:
    ceviri_katologu = {
        "versiyon": {
            "Eski Govde": {"M0": ""},
            "Premium Govde": {"MN": "PREMIUM"},
            "PM Govde": {"MP": "MIKNATISLI"}
        },
        "verimlilik": {
            "IE1": {"1": "IE1"}, "IE2": {"2": "IE2"}, "IE3": {"3": "IE3"},
            "IE4": {"4": "IE4"}, "IE5": {"5": "IE5"},
        },
        "M": {
            "063": {"063": "063"}, "071": {"071": "071"}, "080": {"080": "080"},
            "090": {"090": "090"}, "100": {"100": "100"}, "112": {"112": "112"},
            "132": {"132": "132"}, "160": {"160": "160"}, "180": {"180": "180"},
            "200": {"200": "200"}, "225": {"225": "225"}, "250": {"250": "250"},
            "280": {"280": "280"}, "315": {"315": "315"}, "355": {"355": "355"},
            "400": {"400": "400"}, "450": {"450": "450"}, "500": {"500": "500"},
            "630": {"630": "630"}
        },
        "govde_boyu": {
            "S": {"S": "S"}, "M": {"M": "M"}, "L": {"L": "L"}
        },
        "kutup": {
            "2": {"020": "2"},      "4": {"040": "4"},
            "6": {"060": "6"},      "8": {"080": "8"},
            "10": {"355": "10"},    "12": {"120": "12"},
            "2/4": {"024": "24"},   "4/2": {"042": "42"},
            "4/8": {"048": "48"},   "6/4": {"064": "64"},
            "6/8": {"068": "68"},   "8/2": {"082": "82"},
            "8/4": {"084": "84"},   "8/6": {"086": "86"},
            "2/12": {"212": "212"}, "4/12": {"412": "412"},
            "12/4": {"124": "124"}, "4/16": {"416": "416"}
        },
        "paket_boyu": {
            "A": {"A": "A"},    "B": {"B": "B"},
            "C": {"C": "C"},    "D": {"D": "D"},
            "E": {"E": "E"},    "F": {"F": "F"},
            "K": {"K": "K"},    "P": {"P": "P"},
            "T": {"T": "T"},    "X": {"X": "X"},
            "Y": {"Y": "Y"},    "Q": {"Q": "Q"},
            "RH": {"RH": "RH"}, "A-V": {"A-V": "A-V"}
        },
        "yapi_sekli": {
            "B3": {"0": "B3"}, "B5": {"1": "B5"}, "B9": {"2": "B9"},
            "B14": {"3": "B14"}, "B34": {"4": "B34"}, "B35":{"5":"B35"}
        },
        "flans_olcusu": {
            "YOK": {"00": ""},
            "C90": {"00": "C90"}, "C105": {"01": "C105"}, "C120": {"02": "C120"},
            "C140": {"03": "C140"}, "C160": {"04": "C160"}, "C200": {"05": "C200"},
            "C250": {"06": "C250"}, "A140": {"07": "A140"}, "A160": {"08": "A160"},
            "A200": {"09": "A200"}, "A250": {"10": "A250"}, "A300": {"11": "A300"},
            "A350": {"12": "A350"}, "A400": {"13": "A400"}, "A450": {"14": "A450"},
            "A550": {"15": "A550"}, "A660": {"16": "A660"}, "A800": {"18": "A800"},
            "A1000": {"19": "A1000"}, "A1150": {"20": "A1150"}, "A1370": {"21": "A1370"}
        },
        "govde_malzemesi": {
            "GOVDESIZ": {"0": "GOVDESIZ MOTOR"},
            "ALUMINYUM": {"1": "ALUMINYUM GOVDELI MOTOR"},
            "PIK": {"2": "PIK GOVDELI MOTOR"},
            "CELIK": {"3": "CELIK GOVDELI MOTOR"}
        }
    }

    def __init__(self, kategori_adi, secilen_deger):
        self.kategori_adi = kategori_adi
        self.secilen_deger = secilen_deger
        self.kisa_kod = ""
        self.uzun_aciklamasi = ""
        bulunan_kutu = motor_ozellikleri.ceviri_katologu[self.kategori_adi][self.secilen_deger]
        self.kisa_kod = list(bulunan_kutu.keys())[0]
        self.uzun_aciklamasi = list(bulunan_kutu.values())[0]


class motor:
    def __init__(self, versiyon, verimlilik, M, govde_boyu, kutup, paket_boyu, yapi_sekli,
                 flans_olcusu, govde_malzemesi, resim_no, sira_no, ek_aciklama,
                 ip_sinifi, ozel_frekans, ozel_gerilim, paslanmaz_civata,
                 boya_durum, boya_deger, ozel_kw_durum, ozel_kw_deger, ozel_devir_durum, ozel_devir_deger,
                 standart, isitici, marine, guduk, termistor, vantilator, termik, paslanmaz_rakor, tropikalizeli):

        self.versiyon = motor_ozellikleri("versiyon", versiyon)
        self.verimlilik = motor_ozellikleri("verimlilik", verimlilik)
        self.M = motor_ozellikleri("M", M)
        self.govde_boyu = motor_ozellikleri("govde_boyu", govde_boyu)
        self.kutup = motor_ozellikleri("kutup", kutup)
        self.paket_boyu = motor_ozellikleri("paket_boyu", paket_boyu)
        self.yapi_sekli = motor_ozellikleri("yapi_sekli", yapi_sekli)
        self.flans_olcusu = motor_ozellikleri("flans_olcusu", flans_olcusu)
        self.govde_malzemesi = motor_ozellikleri("govde_malzemesi", govde_malzemesi)
        self.resim_no_deger = str(resim_no).strip() if resim_no else "0000"

        self.sira_no = str(sira_no).zfill(4) if sira_no else ""
        self.ek_aciklama = ek_aciklama
        self.ip_sinifi = ip_sinifi
        self.ozel_frekans = ozel_frekans
        self.ozel_gerilim = ozel_gerilim
        self.paslanmaz_civata = paslanmaz_civata
        self.boya_durum = boya_durum
        self.boya_deger = boya_deger
        self.ozel_kw_durum = ozel_kw_durum
        self.ozel_kw_deger = ozel_kw_deger
        self.ozel_devir_durum = ozel_devir_durum
        self.ozel_devir_deger = ozel_devir_deger
        self.standart = standart
        self.isitici = isitici
        self.marine = marine
        self.guduk = guduk
        self.termistor = termistor
        self.vantilator = vantilator
        self.termik = termik
        self.paslanmaz_rakor = paslanmaz_rakor
        self.tropikalizeli = tropikalizeli

    def base_id_uret(self):
        return (f'{self.versiyon.kisa_kod}{self.verimlilik.kisa_kod}{self.govde_malzemesi.kisa_kod}-'
                f'{self.M.kisa_kod}{self.govde_boyu.kisa_kod}{self.kutup.kisa_kod}{self.paket_boyu.kisa_kod}'
                f'{self.yapi_sekli.kisa_kod}{self.flans_olcusu.kisa_kod}')

    def tip_id_uret(self):
        return (f'{self.base_id_uret()}{self.sira_no}-{self.resim_no_deger}')

    def aciklama_uret(self):
        aciklama = f"{self.verimlilik.uzun_aciklamasi} {self.M.secilen_deger} {self.govde_boyu.kisa_kod}"

        kutup_kisa = self.kutup.secilen_deger  # Arayüzde ne seçildiyse aynen alınır.

        # 4A4A hatası düzeltildi: Bu satır iki kere yazılmıştı, teke düşürüldü.
        aciklama += f"{kutup_kisa}{self.paket_boyu.kisa_kod}"

        if self.vantilator: aciklama += "-V"
        aciklama += f" {self.yapi_sekli.secilen_deger}"

        if self.flans_olcusu.secilen_deger != "YOK":
            aciklama += f" {self.flans_olcusu.secilen_deger}"

        if self.resim_no_deger != "0000":
            aciklama += f" {self.resim_no_deger}"

        if self.ozel_gerilim: aciklama += f" {self.ozel_gerilim}"
        if self.ozel_frekans != "50":
            if self.ozel_frekans: aciklama += f" {self.ozel_frekans}Hz"
        if self.boya_durum and self.boya_deger: aciklama += f" RAL {self.boya_deger} BOYALI"
        if self.ip_sinifi and self.ip_sinifi != "55": aciklama += f" IP{self.ip_sinifi}"
        if self.marine: aciklama += " MARINE"
        if self.guduk: aciklama += " GUDUK"
        if self.termistor: aciklama += " TERMISTORLU"
        if self.isitici: aciklama += " ISITICILI"
        if self.tropikalizeli: aciklama += " TROPIKALIZELI"
        if self.termik: aciklama += " TERMIKLI"
        if self.paslanmaz_rakor: aciklama += " PASLANMAZ RAKORLU"
        if self.paslanmaz_civata != "INOX":
            if self.paslanmaz_civata: aciklama += f" PASLANMAZ CIVATALI ({self.paslanmaz_civata})"

        if self.ek_aciklama:
            temiz_ek = self.ek_aciklama
            if temiz_ek:
                aciklama += f" {temiz_ek}"

        if self.versiyon.secilen_deger == "PM Govde":
            aciklama += " MIKNATISLI"

        aciklama += f" {self.govde_malzemesi.uzun_aciklamasi}"

        # --- KW VE DEVİR MANTIĞI ÖNE ÇEKİLDİ ---
        # Önce Özel KW/Devir var mı ona bakıyoruz, yoksa Standart KW/Devir çekiyoruz
        if self.ozel_kw_durum or self.ozel_devir_durum:
            kw_metni = f" - {self.ozel_kw_deger} kW" if self.ozel_kw_durum and self.ozel_kw_deger else ""
            devir_metni = f" {self.ozel_devir_deger} d/dk" if self.ozel_devir_durum and self.ozel_devir_deger else ""
            aciklama += kw_metni + devir_metni
        else:
            sifre = f"{self.M.secilen_deger}{self.govde_boyu.secilen_deger}{self.kutup.secilen_deger}{self.paket_boyu.secilen_deger}"
            katalog = katalog_oku()
            if sifre in katalog:
                # Kullanıcının istediği format: " - 55 kW 1500 d/dk"
                aciklama += f" - {katalog[sifre]['kw']} kW {katalog[sifre]['devir']} d/dk"
            else:
                aciklama += " - (STANDART KW/DEVİR EKSİK!)"

        # --- PREMIUM EN SONA EKLENİR ---
        if self.versiyon.secilen_deger == "Premium Govde":
            aciklama += " - PREMIUM"

        # Fazla boşlukları temizleyip tek boşluğa indirger
        aciklama = " ".join(aciklama.split())

        return aciklama


def turkce_karakter_temizleme(metin):
    """
    Türkçe karakterleri ASCII karşılıklarına dönüştürür ve büyük harfe çevirir.
    Python'un .upper() metodu 'i' → 'İ' yaptığı için, önce Türkçe harfleri
    manuel olarak dönüştürüp SONRA .upper() çekiyoruz. Bu 'I'/'İ' çakışmasını çözer.
    Hem kullanıcı girdilerinde hem Excel'den okunan verilerle eşleşmede kullanılır.
    """
    if not isinstance(metin, str):
        metin = str(metin)
    # 1. Önce küçük Türkçe harfleri dönüştür (Python upper() 'i'→'İ' yapar, bunu önle)
    yeni_metin = (metin
                  .replace("ı", "i")
                  .replace("ğ", "g")
                  .replace("ü", "u")
                  .replace("ş", "s")
                  .replace("ö", "o")
                  .replace("ç", "c"))
    # 2. Büyük Türkçe harfleri dönüştür
    yeni_metin = (yeni_metin
                  .replace("İ", "I")
                  .replace("Ğ", "G")
                  .replace("Ş", "S")
                  .replace("Ü", "U")
                  .replace("Ö", "O")
                  .replace("Ç", "C"))
    # 3. En son .upper() — artık sadece ASCII büyük harf dönüşümü yapıyor
    return yeni_metin.upper().strip()