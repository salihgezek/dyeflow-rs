# DyeFlow RS v47 Regression Test Checklist

Bu dosya her yeni sürümden önce kontrol edilecek sabit test listesidir. Amaç: bir alan düzeltilirken daha önce düzeltilmiş başka alanların bozulmasını önlemek.

## Stabil Temel
- Stabil ana sürüm: v46
- v47 amacı: stabil temel üzerine test kontrol listesi, sürüm kilidi ve regresyon kontrol sistemi eklemek
- AI Assisted Optimization bu sürüme dahil değildir.

## Zorunlu Manuel Testler

### 1. Açılış ve Login
- [ ] `DyeFlow_RS_ONE_CLICK_START.bat` ile uygulama açılıyor.
- [ ] İlk kullanıcı register olabiliyor.
- [ ] Login sonrası ana ekran düzgün görünüyor.
- [ ] CSS bozulmadan premium arayüz yükleniyor.

### 2. Grafik Motoru
- [ ] Program içindeki grafik doğru çiziliyor.
- [ ] Sıcaklık çizgisi kırmızı.
- [ ] Drain oku aşağı doğru.
- [ ] Drain sonrası bağlantı çizgisi yok.
- [ ] Kimyasal okları ve etiketleri doğru hizalanıyor.
- [ ] Overflow süresinde ±10°C sık sinüs dalgası görünüyor.

### 3. PowerPoint / PDF / Compare Grafik
- [ ] Tek proje PPT grafiği programdaki grafikle aynı görünüyor.
- [ ] Compare PPT grafikleri programdaki grafikle aynı görünüyor.
- [ ] Grafik export eski renderer'a dönmüyor.
- [ ] PowerPoint oluşturma hatasız tamamlanıyor.

### 4. Overflow Hesap Mantığı
- [ ] Overflow suyu water hesabına ekleniyor.
- [ ] Overflow suyu wastewater hesabına ekleniyor.
- [ ] Overflow süresi elektrik aktif süresine ekleniyor.
- [ ] Overflow süresi işçilik süresine ekleniyor.
- [ ] Overflow suyu heating hesabına dahil edilmiyor.
- [ ] Overflow suyu chemical hesabına dahil edilmiyor.

### 5. Compare Dashboard
- [ ] Cost compare batch ve kg başı değerleri doğru gösteriyor.
- [ ] Carbon footprint compare batch için kg CO₂ gösteriyor.
- [ ] Carbon footprint compare kg başına g CO₂/kg gösteriyor.
- [ ] Electricity / Heating / Total CO₂ ayrımı doğru.

### 6. Kimyasal Legend
- [ ] İlk kimyasal ile sonraki kimyasal satırlarının fontları aynı.
- [ ] Legend satırları taşmıyor.
- [ ] Compare ekranında iki projenin kimyasalları ayrı ayrı görünüyor.

## Yeni Değişiklik Yaparken Kural
Yeni bir geliştirme yapılırken sadece ilgili modül değiştirilecek.

Örnek:
- Overflow hesabı değişiyorsa PPT export koduna dokunulmaz.
- PPT grafik değişiyorsa hesaplama motoruna dokunulmaz.
- Login değişiyorsa grafik ve rapor motoruna dokunulmaz.

Her değişiklikten sonra bu checklist yeniden çalıştırılır.
