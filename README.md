# Kervan Market - Mal Kabul v0.1

İlk çalışan prototip.

## Özellikler
- AKINSOFT Excel/CSV ürün listesini içeri alma
- Kamera ile EAN/UPC/Code128/Code39 barkod okuma
- Barkoddan ürün bulma
- Adet / koli / kutu birimi
- Koli/kutu içi adet tanımlama
- Excel'de birim boşsa personelin tanımını koruma
- AKINSOFT listesinde olmayan barkodu uygulama içinden tanımlama
- Mal kabul satırlarını toplama
- Mal kabul geçmişini telefondaki SQLite veritabanında saklama
- İnternet olmadan çalışma

## Beklenen Excel sütunları
En az:
- Barkod
- Ürün Adı

Opsiyonel:
- Ürün Kodu / Stok Kodu
- Birim
- Koli İçi / Kutu İçi / Çevrim

## Çalıştırma
1. Node.js kurulu bilgisayarda klasörü açın.
2. `npm install`
3. `npx expo start`
4. Android cihazda Expo Go ile QR kodu okutun.

## APK için
EAS Build veya yerel Android build aşaması eklenecek.

## Sonraki sürüm
- Tedarikçi seçimi
- Fatura/irsaliye no
- Personel girişi/PIN
- Mal kabul detay sayfası
- Excel dışa aktarma
- Yönetici onayı ve yetki sistemi
- Merkezi Supabase senkronizasyonu
- Aynı üründe farklı koli çevrimi yönetimi
