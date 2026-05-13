# DyeFlow RS Change Control Rules

## Amaç
Her yeni sürümde daha önce düzeltilmiş bölümlerin bozulmasını önlemek.

## Temel Kural
Bir hata düzeltmesi veya yeni özellik sadece kendi modülünde yapılır.

## Korunan Modüller

### 1. Calculation Engine
Dosya: `main.py`
Ana fonksiyon: `calc(project)`
Korunan mantıklar:
- Overflow suyu heating hesabına girmez.
- Overflow suyu chemical hesabına girmez.
- Overflow suyu water/wastewater hesabına girer.
- Overflow süresi electricity/labour süresine girer.
- Carbon footprint batch kg CO₂, kg başına g CO₂/kg olarak raporlanır.

### 2. Chart Engine
Dosyalar: `static/app.js`, `main.py`
Korunan mantıklar:
- Program grafiği ana referanstır.
- PPT/Compare/PDF grafikleri programdaki gerçek chart görünümünden alınır.
- Eski ayrı grafik renderer'a geri dönülmez.

### 3. UI / CSS
Dosyalar: `index.html`, `static/styles.css`
Korunan mantıklar:
- Premium shell korunur.
- Login ekranında CSS yüklenmelidir.
- Sol menü, üst bar ve header yapısı bozulmaz.

### 4. Export Engine
Dosya: `main.py`
Korunan mantıklar:
- Single PPT: 4 slayt yapısı korunur.
- Compare PPT: 4 slayt yapısı korunur.
- Chart export gerçek program grafiğiyle uyumlu kalır.

## Yeni Sürüm Kuralı
Her yeni sürümde:
1. `VERSION_MANIFEST.json` güncellenir.
2. `REGRESSION_TEST_CHECKLIST.md` kontrol edilir.
3. `RUN_REGRESSION_TESTS.bat` çalıştırılır.
4. Hata yoksa ZIP oluşturulur.
