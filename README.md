# AEMOT Vertex Yönetim Paneli
## Akıllı Üretim Planlama, Veri Tutarlılığı ve Yapay Zekâ Destekli Ürün Ağacı Sistemi

AEMOT Vertex, AEMOT Elektrik Motorları'nın üretim planlama süreçlerinde kullanılan motor ve bileşen bilgilerinin yönetimini kolaylaştırmak, geçmiş Excel kayıtlarındaki tutarsızlıkları gidermek ve ürün ağacı oluşturma süreçlerini otomatikleştirmek amacıyla geliştirdiğim Python tabanlı bir masaüstü uygulamasıdır.

Stajım süresince özellikle insan kaynaklı veri hatalarının azaltılması, aynı özelliklere sahip parçaların tutarlı sıra numaralarıyla eşleştirilmesi, ortak ağ üzerindeki Excel dosyalarının güvenilir biçimde yönetilmesi ve makine öğrenmesi destekli ürün ağacı oluşturulması üzerine çalıştım.

Proje; modüler masaüstü arayüzü, kural tabanlı veri doğrulama mekanizmaları, kullanıcı yetkilendirme sistemi, gelişmiş arama özellikleri ve Random Forest tabanlı tahmin modülünü tek bir uygulamada bir araya getirir.

---

## 🎯 Projenin Amacı

Üretim süreçlerinde kullanılan Excel kayıtlarında zamanla oluşan sıra numarası tutarsızlıklarını, farklı yazım biçimlerinden kaynaklanan eşleştirme hatalarını ve manuel veri girişinden doğan sorunları azaltmak hedeflendi.

Bu doğrultuda:
* Motor ve bileşen kodlarının, açıklamalarının ve sıra numaralarının tutarlı şekilde oluşturulması,
* Aynı özellikleri ifade eden farklı metinlerin doğru biçimde eşleştirilmesi,
* Geçmiş kayıtların analiz edilerek uygun sıra numaralarının kullanıcıya önerilmesi,
* Üretim kurallarının ve ölçü standartlarının merkezi olarak yönetilmesi,
* Ortak ağ üzerinde Excel dosyalarına güvenilir veri aktarımı yapılması,
* Motor ve alt bileşenlerden oluşan ürün ağaçlarının otomatik oluşturulması amaçlandı.

---

## ✨ Temel Özellikler

### 1. Akıllı Sıra Numarası Analizi
Projenin en önemli bileşenlerinden biri, geçmiş Excel kayıtlarında aynı özelliklere sahip parçalar için farklı sıra numaralarının kullanılmasından kaynaklanan tutarsızlıkları gidermeye yönelik geliştirdiğim eşleştirme sistemidir.

Örneğin aşağıdaki iki açıklama aynı özellikleri ifade edebilir:
* `IMAK CEKDİRME delikSIZ`
* `CEKDİRME DELİKSİZ IMAK`

Geleneksel metin karşılaştırmasında kelimelerin sırası, büyük-küçük harf farklılıkları ve Türkçe karakterler eşleşmeyi etkileyebilir. Geliştirdiğim yaklaşımda ise metinler normalleştirilerek kelime sırasından bağımsız biçimde karşılaştırılır.

**Uygulanan yöntemler:**
* Büyük-küçük harf farklılıklarının ortadan kaldırılması.
* Türkçe karakter farklılıklarının eşleştirme mantığında normalize edilmesi.
* Kelime sırasından bağımsız özellik karşılaştırması.
* Aynı özellikleri taşıyan geçmiş kayıtların analiz edilmesi.
* Eşleşen kayıtların sıra numaralarının ve kullanım bilgilerinin incelenmesi.

Bu sayede açıklamaların yalnızca yazılış biçimi farklı olduğu için farklı parça kayıtları gibi değerlendirilmesi önlenerek sıra numarası tutarlılığının artırılması hedeflendi.

### 2. İstatistiksel Sıra Numarası Önerisi
Geçmiş üretim kayıtları analiz edilerek eşleşen parçalarda kullanılan sıra numaraları incelenir. Kullanıcıya mevcut kayıtlar ve kullanım istatistikleri üzerinden uygun sıra numarası seçenekleri sunulur.
* Geçmiş kayıtların karşılaştırılması.
* Sıra numarası çakışmalarının tespit edilmesi.
* Kullanım sıklıklarının incelenmesi.
* Uygun sıra numarasının seçilebilmesi.
* Yetkiye bağlı manuel sıra numarası girişi.

Amaç, geçmiş kayıtlarla uyumlu seçimleri kolaylaştırmak ve aynı özellikteki parçalar için gereksiz sıra numarası farklılıklarını azaltmaktır.

### 3. Modüler Motor ve Bileşen Yönetimi
Uygulama, ana motor ve farklı motor bileşenlerinin yönetilebildiği modüler bir masaüstü mimarisine sahiptir. Her bileşenin kendine özgü giriş alanları, doğrulama mekanizmaları ve üretim kuralları bulunur. Ortak işlemlerin yeniden kullanılabilir biçimde tasarlanması, yeni özelliklerin sisteme eklenmesini ve mevcut modüllerin bakımını kolaylaştırır.

### 4. Gelişmiş Arama ve Otomatik Tamamlama
Kullanıcıların çok sayıda özellik ve kayıt arasında daha hızlı işlem yapabilmesi için arama ve kelime tamamlama özellikleri geliştirildi. Örneğin, bir özellik alanına `ima` yazıldığında `IMAK` gibi eşleşen seçeneklerin hızlıca bulunabilmesi sağlandı.

### 5. Merkezi Kural ve Ölçü Yönetimi
Sistemin üretim standartlarının korunması için ölçü ve kural yönetimi mekanizmaları geliştirildi. Yetkili kullanıcılar, sahip oldukları izinler doğrultusunda yeni standart değerler ve kurallar tanımlayabilir. Uygunsuz seçimlerde uyarı mekanizmaları devreye girerek hatalı veri girişlerinin önüne geçilmesine yardımcı olur.

### 6. Ortak Ağ ve Excel Dosya Yönetimi
Uygulamanın farklı bilgisayarlardan ortak Excel dosyalarıyla çalışabilmesi için dosya yolu yönetimi ve veri aktarım mekanizmaları geliştirildi. 
* Kullanıcı, ayarlar üzerinden lokal bir dosyayı veya ağ üzerindeki hedef Excel dosyasını seçebilir.
* Dosyanın başka bir bilgisayarda açık olması nedeniyle oluşabilecek erişim sorunlarına karşı kayıt işlemlerinin yönetilmesine yönelik bir yapı kuruldu. Dosyaya yazma işlemi mümkün olmadığında aktarımın bekletilmesi ve dosya yeniden erişilebilir olduğunda güncellemenin gerçekleştirilmesi sağlandı.

### 7. Rol Tabanlı Kullanıcı Yönetimi ve Loglama
Uygulamada kullanıcıların yetkilerine göre işlem yapabilmesini sağlayan bir yönetim sistemi geliştirildi:
* **Sistem Admini:** Kullanıcı yönetimi, kritik ayarlar, log görüntüleme ve yetkili işlemleri.
* **Yetkili:** İzin verilen kural ve ölçü yönetimi işlemleri ile ilgili kritik işlemler.
* **Standart Kullanıcı:** Yetkisi dahilindeki veri giriş ve sorgulama işlemleri.
* **Loglama:** Gerçekleştirilen kritik işlemler log mekanizmasıyla kayıt altına alınarak geriye dönük inceleme ve işlem takibi kolaylaştırıldı.

### 8. Yapay Zekâ Destekli Ürün Ağacı (BOM)
Projenin makine öğrenmesi modülünde geçmiş üretim reçetelerinden yararlanılarak yeni motorlara ait alt bileşenleri tahmin eden bir sistem geliştirildi.

**Model Geliştirme Süreci:**
1. Geçmiş üretim verileri derlendi ve eğitim için hazırlandı.
2. Kategorik veriler `LabelEncoder` ile sayısal değerlere dönüştürüldü.
3. `Scikit-learn` kullanılarak **Random Forest** modeli eğitildi.
4. Eğitilen model `.pkl` formatında uygulamaya entegre edildi.
5. Kullanıcının girdiği motor kodu ve özelliklerinden yararlanılarak alt parçalık ürün ağacı tahmin edildi.
6. Tahmin edilen bileşenler tablo üzerinde gösterilerek ilgili parça modüllerine aktarılabilir hale getirildi.

> ⚠️ **Model Performansı:** Geliştirme ve test sürecinde yaklaşık **%75–80** doğruluk aralığı elde edildi. Bu sonuç kullanılan veri seti ve test koşullarıyla sınırlıdır. Özellikle eğitim verilerinde yeterince temsil edilmeyen motor türlerinde tahminlerin uzman kontrolünden geçirilmesi gerekir.

### 9. Ürün Ağacından Bileşen Kayıtlarına Otomatik Aktarım
Ürün ağacı modülü, yalnızca tahmin sonuçlarını göstermekle kalmayıp bu sonuçların uygulamanın ilgili bileşen modüllerine aktarılmasını da destekler. Motor kodu ve özelliklerine göre oluşturulan ürün ağacı üzerinden bileşen seviyeleri belirlenir; ilgili parçaların kod ve açıklama bilgilerinin oluşturulması ve Excel sayfalarına aktarılması sağlanır.

---

## 🛠️ Kullanılan Teknolojiler

| Teknoloji | Kullanım Amacı |
| :--- | :--- |
| **Python** | Uygulama geliştirme ve iş mantığı |
| **PyQt6** | Masaüstü kullanıcı arayüzü |
| **Pandas** | Veri işleme, filtreleme ve analiz |
| **Scikit-learn** | Makine öğrenmesi |
| **Random Forest** | Ürün ağacı tahmini |
| **LabelEncoder** | Kategorik verileri sayısal değerlere dönüştürme |
| **JSON** | Kullanıcı ve yapılandırma verileri |
| **Excel / xlwings** | Üretim kayıtları ve veri yönetimi |
| **PyInstaller** | Çalıştırılabilir uygulama oluşturma |

---

## 🧠 Bu Projede Kazandığım Deneyimler

* Gerçek üretim verilerindeki tutarsızlıkları analiz etme ve çözüm geliştirme.
* Normalleştirme ve kelime sırasından bağımsız metin eşleştirme algoritmaları kurma.
* Modüler Python ve PyQt6 ile kullanıcı odaklı masaüstü arayüzleri geliştirme.
* Makine öğrenmesi veri seti hazırlama, model eğitimi ve entegrasyonu.
* Excel tabanlı ortak ağ (network) dosya erişim çakışmalarını yönetme mekanizmaları kurma.
* Rol tabanlı yetkilendirme ve detaylı işlem loglama altyapısı tasarlama.

---

## 📌 Proje Bilgileri

* **Proje Adı:** AEMOT Vertex Yönetim Paneli
* **Geliştirme Alanları:** Masaüstü Yazılım, Veri İşleme, Makine Öğrenmesi
* **Programlama Dili:** Python
* **Arayüz Teknolojisi:** PyQt6
* **Makine Öğrenmesi:** Random Forest
* **Staj Dönemi:** 06.07.2026 – 17.08.2026
* **Çalışma Süresi:** 30 İş Günü

---

## 🔗 Kaynak Kod
* [AEMOT Vertex Yönetim Paneli – GitHub Deposu](https://github.com/KULLANICI_ADIN/REPO_ADIN) *(Düzenlenecek)*

---

> **⚠️ Not:** Proje, gerçek üretim planlama ihtiyaçlarına yönelik geliştirilmiştir. Kurumsal veri gizliliği nedeniyle şirketin özel verileri ve üretim kayıtları herkese açık biçimde paylaşılmamıştır. Makine öğrenmesi çıktıları karar destek amacı taşır ve üretim süreçlerinde kullanılmadan önce kontrol edilmelidir.

**Geliştiren:** Muhammet Ali Aslanhan
