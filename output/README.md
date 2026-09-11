# Materi Keselamatan Twenty Swim (versi orang tua & anak)

Semua file HTML berdiri sendiri (CSS di dalam file), memuat font Montserrat dan Lora dari Google Fonts dengan fallback sans-serif/serif.

| File | Format | Isi |
|---|---|---|
| `lembar-persetujuan-orangtua.html` | A4 portrait, 1 halaman | Lembar persetujuan formal, 10 aturan lengkap, kolom tanda tangan |
| `poster-orangtua-A4.html` | A4 portrait, 1 halaman | Poster "10 Aturan Aman untuk Ayah & Bunda" |
| `kartu-anak-A5.html` | A5 portrait, 1 halaman | Kartu "Janji Perenang Pintar" (7 janji, kotak centang) |
| `carousel-ig-orangtua/slide-1..6.html` | 1080 x 1350 px | 6 slide carousel Instagram |
| `carousel-ig-orangtua/caption.txt` | teks | Caption Instagram + 5 tagar |
| `whatsapp-onboarding.txt` | teks | 3 pesan siap kirim untuk admin |

## 1. Menyematkan logo asli

Logo belum tersedia sebagai file di repo ini, sehingga setiap HTML memakai wordmark sementara (SVG) di antara penanda `<!-- LOGO:START -->` dan `<!-- LOGO:END -->`.

```
python3 output/tools/sematkan-logo.py logo-twentyswim.png
```

Skrip mengganti wordmark sementara dengan `<img src="data:image/png;base64,...">` di semua HTML. Tidak ada CSS filter yang diterapkan pada logo; di latar navy logo diletakkan di dalam kotak putih.

## 2. Mencetak

Buka file di Chrome/Edge, lalu Cetak (Ctrl+P): ukuran kertas sesuai (A4 atau A5), margin "Default" (file sudah mengatur 12 mm), aktifkan "Background graphics". Simpan sebagai PDF bila perlu.

Kartu anak 2 per lembar: pilih kertas A4, orientasi landscape, opsi "Pages per sheet: 2".

## 3. Mengekspor slide Instagram ke PNG

Cara manual: buka `slide-N.html` di Chrome, buka DevTools (F12), aktifkan mode perangkat (Ctrl+Shift+M), atur ukuran 1080 x 1350, lalu klik menu tiga titik pada bilah perangkat dan pilih "Capture screenshot".

Cara otomatis (perlu Node.js dan Playwright):

```
npm install playwright
npx playwright install chromium
node output/tools/render-slides.js
```

Hasil PNG tersimpan di `output/carousel-ig-orangtua/png/`.
