# Materi Keselamatan Twenty Swim (versi orang tua & anak)

Semua file HTML berdiri sendiri (CSS di dalam file), memuat font Montserrat dan Lora dari Google Fonts dengan fallback sans-serif/serif.

| File | Format | Isi |
|---|---|---|
| `lembar-persetujuan-orangtua.html` | A4 portrait, 1 halaman | Lembar persetujuan formal, 11 aturan lengkap, kolom tanda tangan |
| `poster-orangtua-A4.html` | A4 portrait, 1 halaman | Poster "11 Aturan Aman untuk Ayah & Bunda" |
| `kartu-anak-A5.html` | A5 portrait, 1 halaman | Kartu "Janji Perenang Pintar" (8 janji, kotak centang) |
| `carousel-ig-orangtua/slide-1..6.html` | 1080 x 1350 px | 6 slide carousel Instagram |
| `carousel-ig-orangtua/caption.txt` | teks | Caption Instagram + 5 tagar |
| `whatsapp-onboarding.txt` | teks | 3 pesan siap kirim untuk admin |
| `halaman-orangtua/index.html` | halaman web, 1 file | Halaman beranimasi 11 aturan untuk orang tua, siap di-hosting di Netlify |
| `film-orangtua/film-keselamatan-anak-twentyswim.mp4` | video MP4, 1080x1920, 48 detik | Film pendek animasi untuk WhatsApp Status / Instagram Reels |
| `film-orangtua/film.html` | halaman animasi, 1 file | Sumber film (garis waktu animasi); dapat dirender ulang jadi MP4 |

## 1. Menyematkan logo asli

Logo asli (`logo-twentyswim.png`, 800 x 391 px, PNG transparan) sudah tersemat sebagai base64 di semua HTML, di antara penanda `<!-- LOGO:START -->` dan `<!-- LOGO:END -->`. Bila logo diperbarui, jalankan ulang:

```
python3 output/tools/sematkan-logo.py logo-twentyswim.png
```

Skrip mengganti isi di antara penanda dengan `<img src="data:image/png;base64,...">` di semua HTML. Tidak ada CSS filter yang diterapkan pada logo; di latar navy logo diletakkan di dalam kotak putih.

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

## 4. Halaman web untuk orang tua (Netlify)

`output/halaman-orangtua/index.html` berdiri sendiri (logo base64, CSS dan JS inline, font dari Google Fonts).

Deploy drag-drop: buka https://app.netlify.com/drop, seret folder `output/halaman-orangtua` ke area unggah, tunggu tautan `*.netlify.app` muncul. Untuk memperbarui, buka situs itu di dasbor Netlify, tab Deploys, lalu seret folder yang sama lagi.

Nomor WhatsApp admin diisi pada konstanta `NOMOR_WA` di bagian `<script>` paling bawah file (format 62812xxxxxxx). Selama kosong, tombol WhatsApp tidak ditampilkan.

## 5. Film pendek (MP4 untuk WhatsApp / Instagram Reels)

`output/film-orangtua/film-keselamatan-anak-twentyswim.mp4` adalah film vertikal 1080x1920, 48 detik, tanpa suara, dengan karakter manusia beranimasi (Bunda, pelatih, dan anak berbaju cerah) yang bernapas, berkedip, melangkah, memanggil, dan berenang: pembuka brand, alur check-in sampai check-out (siapa yang mengawasi anak di tiap tahap), 3 aturan paling penting (teks persis poster), lalu penutup brand. Ukuran sekitar 1,8 MB, format H.264 yang aman di WhatsApp dan Instagram.

Kirim langsung file MP4-nya. Untuk Instagram Reels atau WhatsApp Status, unggah seperti video biasa; rasio 9:16 sudah pas layar penuh HP.

Merender ulang (jika `film.html` diubah): butuh Node.js, Playwright, dan ffmpeg.

```
FFMPEG=$(python3 -c "import imageio_ffmpeg;print(imageio_ffmpeg.get_ffmpeg_exe())") node output/tools/render-film.js
```

Bila ffmpeg sudah ada di PATH, cukup `node output/tools/render-film.js`. Font Montserrat dan Lora sudah ditanam di `film.html` sebagai base64, jadi render tidak butuh internet.
