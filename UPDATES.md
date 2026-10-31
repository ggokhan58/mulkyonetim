# GitHub güncellemeleri

## 31 Ekim 2026 — yerel kira tahsilatı düzeltmesi

- Güncel kira bedeli ve son ödeme kaydının tahsilatı ayrı gösterilir.
  Tahsilat güncel kira eksi eski kesinti olarak yeniden hesaplanmaz;
  ödeme kaydının `amount_paid` tutarı kullanılır.
- Mülk düzenlendikten sonra son ödeme bilgileri veritabanından yeniden okunur.
- Geçmiş ayda kayıt yoksa ödeme tutarı boş başlar; o dönemde alınan gerçek
  tutar girilmelidir. Yeni döneme önceki ödemenin kesintisi taşınmaz.
- Var olan ödeme kayıtları ve veritabanı şeması değiştirilmez.
- Kontrol: `node --test --test-isolation=none scripts/rent-history.test.cjs scripts/updater.test.cjs`
  (Node.js 24), `npx tsc --noEmit`, `npx vite build`.

GitHub API kontrolünde yayımlanmış son sürüm `1.0.5` (7 Ağustos 2026).
Önceki yerel `1.0.6` paketinin üzerine güncelleme alınabilmesi için bu
düzeltme `1.0.7` sürümü olarak paketlenir. Sürüm notları `RELEASE_NOTES.md` içindedir.

14 Eylül 2026 kontrolünde GitHub 1.0.5 sürümündeki `latest.yml`,
`mulk-yonetim-setup-1.0.5.exe` istiyordu; yüklenen dosyanın adı
`Mulk.Yonetim.Setup.1.0.5.exe` idi. Bu uyuşmazlık indirmeyi engeller.
Yerelde sonradan üretilen 1.0.5 paketinin özeti de GitHub'daki metadata
ile farklıydı; farklı derlemelerin dosyaları karıştırılmamalı.

## Yeni sürüm hazırlama

1. Her dağıtımda package.json ve package-lock.json sürümünü artırın.
2. Windows üzerinde `npm ci` ve `npm run build` çalıştırın.
3. Build sonunda doğrulama kurulum dosyasının adını, boyutunu, SHA512
   özetini ve paket içindeki GitHub adresini kontrol eder.
4. GitHub Releases'ta aynı sürüme ait bir taslak açın (ör. `v1.0.6`).
   `release/1.0.6` klasöründeki şu üç dosyayı adlarını değiştirmeden yükleyin:
   - `mulk-yonetim-setup-1.0.6.exe`
   - `mulk-yonetim-setup-1.0.6.exe.blockmap`
   - `latest.yml`
5. Üç dosya da yüklendikten sonra taslağı normal sürüm olarak yayımlayın.
   Taslak veya ön sürüm kullanıcıların kararlı sürüm güncellemesinde görünmez.

Build komutu yalnızca yerel paket üretir; GitHub'a otomatik yüklemez.
Etkin paket ayarları package.json içindeki `build` alanındadır.
`electron-builder.json5` eski şablondur; `--config electron-builder.json5`
ile paket üretmeyin.

## Kullanıcı tarafı

Kurulu uygulama açılıştan 10 saniye sonra ve saatte bir kontrol eder.
`Güncellemeleri Kontrol Et` düğmesi elle kontrol ve hata bildirimi sağlar.
İndirilen paket için yeniden başlatma seçeneği gösterilir; ertelenirse
normal kapanışta kurulur. Günlük `%APPDATA%/mulk-yonetim/updater.log`
altında tutulur (Electron userData yolu özelleştirilmişse o klasördedir).
Geliştirme modunda indirme yapılmaz.

Kontroller: `npm run test:updater`, `npm run build`.
Gerçek uçtan uca doğrulama için önceki kurulu sürümden, yayımlanmış daha
yüksek sürüme güncelleme denenmelidir. Kaynak kodu GitHub'a göndermek tek
başına kurulu uygulamaları güncellemez.
