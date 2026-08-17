import os
import traceback  # Hata detaylarını yakalamak için eklendi
import xlwings as xw
from arayuz.mantik.json_islemleri import ayar_excel_yolunu_getir, excel_yedekle, \
    hata_logla  # Yedekleme ve log fonksiyonları dahil edildi


# DÜZELTME: Artık min_satir parametrik. Başlık 8'deyse minimum sınır 7 olur.
def gercek_son_satiri_bul_xw(sheet, min_satir=11):
    """Aşağıdan yukarı çıkarak gerçek dolu satırı bulur."""
    son_satir_a = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    son_satir_b = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row

    return max(son_satir_a, son_satir_b, min_satir)


# DÜZELTME: baslik_satiri=12 varsayılan olarak eklendi. Rotorlu mil buraya 8 yollayacak.
def toplu_excele_aktar(tablo_widget, sayfa_indeksi=0, baslik_satiri=12):
    GUNCEL_EXCEL_YOLU = ayar_excel_yolunu_getir()

    if tablo_widget.rowCount() == 0:
        return False, "Tabloda aktarılacak veri yok!"

    if not os.path.exists(GUNCEL_EXCEL_YOLU):
        return False, f"Excel dosyası bulunamadı, ayarlardan dosya yolunu kontrol edin:\n{GUNCEL_EXCEL_YOLU}"

    # ---------------------------------------------------------
    # İŞTE BURASI: Excel'e veri basmadan hemen önce yedeğini alıyoruz!
    excel_yedekle(GUNCEL_EXCEL_YOLU)
    # ---------------------------------------------------------

    toplam_sutun = tablo_widget.columnCount()
    son_sutun_harfi = chr(64 + toplam_sutun) if toplam_sutun <= 26 else 'Z'

    yeni_veriler = []
    for satir in range(tablo_widget.rowCount()):
        satir_verisi = []
        for sutun in range(toplam_sutun):
            item = tablo_widget.item(satir, sutun)
            veri = item.text().strip() if item else ""

            # DÜZELTME: Ana Motor için 12,13; Rotorlu Mil için 5,6. sütunlara tek tırnak koy.
            if sutun in [5, 6, 12, 13]:
                veri = f"'{veri}"

            satir_verisi.append(veri)
        yeni_veriler.append(satir_verisi)

    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(GUNCEL_EXCEL_YOLU)
        ws = wb.sheets[sayfa_indeksi]

        # DÜZELTME: Başlık hücresini gönderilen satıra göre belirliyoruz (Örn: H8 veya P12)
        baslik_hucre = f'{son_sutun_harfi}{baslik_satiri}'
        if ws.range(baslik_hucre).value != "KAYIT TARİHİ":
            ws.range(baslik_hucre).value = "KAYIT TARİHİ"
            ws.range(baslik_hucre).api.Font.Bold = True

        # Minimum satırı başlığın bir üstü olarak belirliyoruz
        son_satir = gercek_son_satiri_bul_xw(ws, min_satir=(baslik_satiri - 1))
        baslama_satiri = son_satir + 1

        ws.range(f'A{baslama_satiri}').value = yeni_veriler

        if son_satir > baslik_satiri:
            kopyalanacak_alan = f'A{son_satir - 1}:{son_sutun_harfi}{son_satir}'
            hedef_alan = f'A{baslama_satiri}:{son_sutun_harfi}{baslama_satiri + len(yeni_veriler) - 1}'

            ws.range(kopyalanacak_alan).copy()
            ws.range(hedef_alan).paste('formats')

            ws.range(hedef_alan).api.HorizontalAlignment = -4108
            ws.range(f'A{baslama_satiri}:A{baslama_satiri + len(yeni_veriler) - 1}').api.HorizontalAlignment = -4131

            if toplam_sutun >= 2:
                ws.range(f'B{baslama_satiri}:B{baslama_satiri + len(yeni_veriler) - 1}').api.HorizontalAlignment = -4131

            if toplam_sutun >= 12:
                ws.range(f'L{baslama_satiri}:L{baslama_satiri + len(yeni_veriler) - 1}').api.HorizontalAlignment = -4131

        wb.save()
        wb.close()
        app.quit()

        return True, f"{len(yeni_veriler)} adet veri {baslama_satiri}. satırdan itibaren başarıyla kaydedildi!"

    except Exception as e:
        if app is not None:
            try:
                app.quit()
            except:
                pass

        # ---------------------------------------------------------
        # HATA LOGLAMA: Çökme olursa hatanın tüm geçmişini metin dosyasına yazdırır
        hata_logla(traceback.format_exc())
        # ---------------------------------------------------------

        return False, f"Aktarım sırasında hata:\n{e}"