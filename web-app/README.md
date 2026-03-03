# Cihaz Fiyat ve QR Web Uygulamasi

Bu uygulama tarayicida calisir. Kurulum gerektirmez.

## Ozellikler
- Satir ekleme ve tum alanlari girme
- Her link alani icin `Link Ekle + QR Uret` butonu
- `.xlsx` dosyasini `Excel Ice Aktar` ile forma otomatik doldurma
- QR kodlarin link hucre konumuna yerlestirilmesi (Excel exportta E ve H sutunlari)
- Cihaz fiyat tablosunu Excel olarak export etme
- Cikti dosyasi mevcut tablo duzeni ile uyumludur (A-J sutunlari)

## Calistirma
Secenek 1:
- `index.html` dosyasini tarayicida ac.

Secenek 2 (onerilen):
```bash
cd '/Users/esreferdogan/Desktop/yazılım/Satın Alma/web-app'
python3 -m http.server 5500
```
Sonra tarayicida `http://localhost:5500` ac.

## Kullanim
1. `Satir Ekle` ile yeni kayit ac.
2. Istersen `Excel Ice Aktar` ile mevcut dosyani yukle.
3. Link alanina URL yaz.
4. `Link Ekle + QR Uret` butonuna bas.
5. Diger alanlari doldur.
6. `Excel Export` ile `.xlsx` dosyasini indir.
