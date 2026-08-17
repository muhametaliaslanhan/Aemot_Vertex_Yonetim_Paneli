import pandas as pd

# 1. Asıl devasa veriyi oku
dosya_yolu = r"C:\Users\aemot.stajyer\Desktop\urun_agaci_verileri\urun_agaci_verileri_bir.xlsx"
df = pd.read_excel(dosya_yolu)

# 2. Tüm benzersiz ana motor (Tepekodu1) listesini al
motor_listesi = df['Tepekodu1'].dropna().unique().tolist()

# 3. Kendimize test için 5 farklı motor seçelim (ilk 5'i alıyoruz, istersen rastgele de alabilirsin)
# Bu motorların B3, B14, B35 gibi farklı tiplerden olması test için iyi olur.
ornek_motorlar = motor_listesi[:5]

# 4. SADECE bu 5 motorun ürün ağacını (tüm alt parçalarını) filtrele
# Senin o mil.py, kapak.py gibi dosyaların testini yapacağın veriler bunlar.
haftasonu_test_df = df[df['Tepekodu1'].isin(ornek_motorlar)]

# 5. Eve götüreceğin ufak ve güvenli Excel dosyasını oluştur
cikis_yolu = r"C:\Users\aemot.stajyer\Desktop\haftasonu_mock_data.xlsx"
haftasonu_test_df.to_excel(cikis_yolu, index=False)

print(f"İşlem Tamam! {len(ornek_motorlar)} adet motorun tüm alt parçaları (Toplam {len(haftasonu_test_df)} satır) dışa aktarıldı.")
print("Bu 'haftasonu_mock_data.xlsx' dosyasını flash belleğine al ve eve git.")