# arayuz/mantik/oturum.py

aktif_kimlik = ""
aktif_ad = ""
aktif_soyad = ""

def tam_isim():
    if aktif_ad and aktif_soyad:
        return f"{aktif_ad.upper()} {aktif_soyad.upper()}"
    return "Bilinmeyen Kullanıcı"