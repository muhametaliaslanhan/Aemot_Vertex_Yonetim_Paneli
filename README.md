# AEMOT Vertex Yönetim Paneli

**Yapay Zekâ Destekli Üretim Planlama ve Ürün Ağacı Yönetim Sistemi**

AEMOT Vertex Yönetim Paneli, elektrik motoru üretim süreçlerinde motor ve parça bilgilerinin yönetilmesi, üretim kurallarının uygulanması, sıra numarası tutarsızlıklarının azaltılması ve ürün ağacı (Bill of Materials - BOM) oluşturulması amacıyla geliştirilen Python tabanlı bir masaüstü uygulamasıdır.

Proje kapsamında modüler bir yönetim paneli, rol tabanlı kullanıcı yetkilendirme sistemi, Excel tabanlı veri yönetimi ve geçmiş üretim verilerinden yararlanarak ürün ağacı tahmini yapan bir makine öğrenmesi modülü geliştirildi.

Bu proje, 6 Temmuz – 17 Ağustos 2026 tarihleri arasındaki 30 iş günlük stajım kapsamında geliştirildi.

## 🚀 Projenin Amacı

Üretim planlama süreçlerinde manuel veri girişinden kaynaklanan hataları azaltmak, motor ve parça bilgilerinin standartlaştırılmasını sağlamak ve tekrarlanan işlemleri otomatikleştirmek hedeflendi.

Projenin temel hedefleri:

- Motor ve alt parçalar için mühendislik kurallarına uygun kod ve açıklamalar oluşturmak.
- Aynı özelliklere sahip parçalar için sıra numarası tutarsızlıklarını tespit etmek.
- Kullanıcı rolleri ve erişim kontrolleriyle sistem yönetimini düzenlemek.
- Ortak ağ üzerindeki Excel dosyalarıyla veri alışverişini yönetmek.
- Makine öğrenmesi ile motorlara ait ürün ağaçlarını tahmin etmek.
- Tekrarlanan veri girişlerini azaltarak üretim planlama sürecini hızlandırmak.

## 🛠️ Kullanılan Teknolojiler

| Teknoloji | Kullanım Alanı |
|---|---|
| Python | Uygulama geliştirme ve iş mantığı |
| PyQt6 | Masaüstü kullanıcı arayüzü |
| Pandas | Veri okuma, işleme ve filtreleme |
| Scikit-learn | Makine öğrenmesi modeli |
| Random Forest | Ürün ağacı tahmini |
| LabelEncoder | Kategorik verilerin sayısal dönüşümü |
| JSON | Kullanıcı, ayar ve yapılandırma verileri |
| Excel | Üretim kayıtları ve veri yönetimi |
| xlwings | Excel dosyalarıyla etkileşim |
| PyInstaller | Uygulamanın `.exe` formatında paketlenmesi |

## ✨ Temel Özellikler

### 1. Modüler Masaüstü Yönetim Paneli

PyQt6 ile geliştirilen uygulama, farklı motor ve parça türlerini yöneten 20 arayüz modülünden oluşur. Ortak işlemler mantık katmanında toplanırken parçaya özgü alanlar ve üretim kuralları ilgili modüllerde uygulanır.

### 2. Kullanıcı Yönetimi ve Yetkilendirme

- Sistem Admini, Yetkili ve Standart olmak üzere üç kullanıcı rolü.
- Kullanıcı rollerine göre menü ve işlem erişimlerinin sınırlandırılması.
- Kullanıcı ve sistem ayarlarının yönetilmesi.
- Kritik işlemlerin kayıt altına alınması.

### 3. Motor DNA ve Üretim Kuralı Sistemi

Motorların gövde, kutup, verimlilik, güç ve opsiyonel özellikleri belirli kurallar doğrultusunda işlenerek standart açıklama metinleri oluşturulur.

- Otomatik açıklama ve kod oluşturma.
- Özel güç ve devir girişleri için doğrulama kontrolleri.
- Parçaya özgü mühendislik kuralları.
- Hatalı veya yasaklı kombinasyonların engellenmesi.
- Kayıtların Excel'e aktarılmadan önce tabloda kontrol edilmesi.

### 4. Sıra Numarası Çakışma Analizi

Geçmiş üretim kayıtlarında aynı özelliklere sahip parçalar için farklı sıra numaralarının kullanılması sorununa yönelik bir analiz mekanizması geliştirildi.

Sistem, parça özelliklerinden oluşturulan DNA metnini geçmiş kayıtlarla karşılaştırarak eşleşen kayıtları ve kullanılan sıra numaralarını belirler. Çakışma durumunda kullanıcıya mevcut numaraları inceleme, seçim yapma veya yetkili olarak manuel numara girme seçenekleri sunulur.

### 5. Excel Veri Yönetimi ve Ağ Optimizasyonu

Ortak ağ üzerindeki Excel dosyalarıyla çalışırken oluşabilecek dosya erişim ve kilitlenme sorunlarını azaltmak amacıyla geçici dosya ve veri aktarım mekanizması geliştirildi.

- Excel verilerinin okunması, filtrelenmesi ve güncellenmesi.
- Toplu Excel aktarımı.
- Kayıtlarla birlikte kullanıcı ve tarih-saat bilgilerinin tutulması.
- Dosyaya erişilemediğinde kullanıcıya hata bildirimi gösterilmesi.
- Geçici dosyalar üzerinden veri hazırlama ve ana dosyaya aktarma yaklaşımı.

### 6. Yapay Zekâ Destekli Ürün Ağacı (BOM)

Projenin makine öğrenmesi bileşeninde geçmiş üretim reçeteleri kullanılarak bir Random Forest modeli eğitildi.

**Model geliştirme süreci:**

1. Geçmiş üretim verileri derlendi ve model eğitimine uygun şekilde düzenlendi.
2. Kategorik veriler `LabelEncoder` kullanılarak sayısal değerlere dönüştürüldü.
3. Scikit-learn ile Random Forest modeli eğitildi.
4. Eğitim sonucunda elde edilen model `.pkl` formatında uygulamaya entegre edildi.
5. Kullanıcının girdiği motor tepekodu ve açıklamasından yararlanılarak 19 alt parçalık ürün ağacı tahmin edildi.
6. Tahmin edilen parçalar tablo üzerinde gösterildi ve ilgili parça modüllerine aktarılabildi.

**Model performansı:** Staj çalışmaları sırasında yapılan testlerde yaklaşık %75–80 doğruluk aralığı elde edildi. Bu değer, kullanılan veri seti ve gerçekleştirilen testlerle sınırlıdır; her motor türü için aynı performansın sağlanacağı anlamına gelmez. Özellikle geçmiş verilerde bulunmayan özel motorlarda tahmin sonuçlarının uzmanlar tarafından kontrol edilmesi gerekir.

### 7. Hata Ayıklama ve Dağıtım

Geliştirme sürecinin sonunda modüller test edildi, tespit edilen hatalar giderildi ve uygulama PyInstaller kullanılarak çalıştırılabilir `.exe` formatında paketlendi. Son aşamada şirketin ortak ağında gerçek kullanım senaryoları üzerinden testler gerçekleştirildi.

## 🧠 Projede Kazanılan Deneyimler

- Modüler Python uygulaması tasarlama ve geliştirme.
- PyQt6 ile masaüstü arayüzleri oluşturma.
- Makine öğrenmesi için veri hazırlama ve kategorik veri dönüştürme.
- Random Forest modeli eğitme ve uygulamaya entegre etme.
- İş kuralları, veri doğrulama ve hata yönetimi geliştirme.
- Excel tabanlı veri işleme ve ağ üzerinde dosya yönetimi.
- Kullanıcı rolleri ve işlem kayıtları oluşturma.
- Yazılımı paketleme, test etme ve gerçek kullanım ortamına hazırlama.

## 📂 Proje Deposu

Kaynak kodlara ve proje dosyalarına GitHub deposu üzerinden ulaşabilirsiniz:

🔗 **[AEMOT Vertex Yönetim Paneli – GitHub](https://github.com/muhametaliaslanhan/Aemot_Vertex_Yonetim_Paneli)**

## ⚠️ Not

Proje, şirketin üretim planlama süreçlerine yönelik geliştirilmiştir. Kurumsal veri gizliliğini korumak amacıyla örneklerde ve ekran görüntülerinde kullanılan veriler sınırlandırılabilir. Makine öğrenmesi modülü karar destek amacı taşır; tahmin edilen ürün ağaçlarının üretimde kullanılmadan önce kontrol edilmesi gerekir.

---

**Geliştirici:** Muhammet Ali Aslanhan  
**Staj Süresi:** 06.07.2026 – 17.08.2026  
**Staj Süresi:** 30 iş günü  
**Alanlar:** Python · Masaüstü Yazılım Geliştirme · Makine Öğrenmesi · Veri İşleme
