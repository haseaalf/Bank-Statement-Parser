# Bank Statement Parser (BSP)

Aplikasi Windows untuk mengubah PDF rekening koran berbagai bank menjadi Excel standar,
lengkap dengan verifikasi saldo dan rekap mutasi bulanan. **Tidak membutuhkan Excel atau WPS.**
File hasil berformat `.xlsx` dan bisa dibuka di Excel, WPS, LibreOffice, atau Google Sheets.

> Branch `feature/bsp-exe` hanya berisi file yang dibutuhkan untuk **memakai** BSP.exe.
> Kode sumber pengembangan (versi Excel, tes, alat bantu) ada di branch `feature/bsp-v1`.

## Instalasi (cukup sekali per PC)

1. Di halaman GitHub, pilih branch **feature/bsp-exe**, lalu klik **Code > Download ZIP**.
2. Klik kanan file ZIP, pilih **Extract All...**, lalu pilih folder tetap, misalnya
   `Documents\Bank Statement Parser`. Jangan menjalankan dari dalam ZIP.
3. Double-click **`Install_BSP.bat`**, lalu tunggu sampai muncul **INSTALASI SELESAI**. Skrip ini:
   - memasang R 4.2.2 bila belum ada (khusus user ini, tanpa hak administrator);
   - memasang paket R yang dibutuhkan;
   - membuat **`BSP.exe`** di folder program, memakai compiler bawaan Windows;
   - membuat shortcut **Bank Statement Parser** di desktop.

**PC tanpa internet:** jalankan `Install_BSP.bat -PrepareOffline` di PC yang online. Perintah ini
mengunduh installer R dan paket R ke `setup\offline\`. Salin seluruh folder ke PC tujuan, lalu
jalankan `Install_BSP.bat` di sana.

## Cara pakai

1. Buka **Bank Statement Parser** dari desktop.
2. Seret file PDF rekening koran ke kotak di atas, atau klik kotak itu untuk memilih file.
   Banyak file, bank, rekening, dan nasabah bisa diproses sekaligus.
3. Opsional: klik **Deteksi Bank** untuk mengecek bank setiap file.
   - Jika status **Tidak dikenali** atau **Ragu**, pilih bank di kolom **Bank**.
   - Jika status **Perlu password**, isi kolom **Password**.
4. Pilih opsi berikut, lalu klik **Proses**:
   - **Mode output**: *Per file* (1 PDF menjadi 1 Excel) atau *Gabung per rekening* (semua bulan satu
     rekening menjadi 1 Excel).
   - **Toleransi saldo**: selisih yang masih dianggap cocok (default 0,01).
   - **Buat rekap mutasi**: satu file rekap berisi sheet REKAP dan satu sheet per rekening (default menyala).
5. Hasil muncul di daftar **Hasil**, semuanya sudah tercentang. Klik **Unduh** per baris, atau
   **Unduh yang Dicentang...** untuk menyimpan sekaligus ke folder pilihan.
6. Baca **Log proses**: hijau berarti OK, oranye berarti perlu dicek, merah berarti gagal.

Tombol **Hapus baris terpilih** dan **Hapus daftar** mengatur daftar file. Di daftar **Hasil**,
kotak centang di judul kolom mencentang atau menghapus centang semua hasil, dan **Hapus yang dicentang**
membuang hasil yang dicentang dari daftar (file yang sudah diunduh tidak terpengaruh).
**Perbarui daftar bank** menampilkan bank yang didukung.

Semua file diproses di komputer ini dan tidak dikirim ke internet. Hasil sementara dihapus saat
aplikasi ditutup, jadi unduh dulu hasil yang diperlukan.

## Bank yang didukung

| Kode | Bank | Format |
|---|---|---|
| BRI | Bank Rakyat Indonesia | Laporan Transaksi Finansial (BRISIM) |
| BCA | Bank Central Asia | e-Statement Rekening Giro |
| BNI | Bank Negara Indonesia | Account Statement |
| MANDIRI | Bank Mandiri | Account Statement (Kopra) |
| MANDIRI | Bank Mandiri | e-Statement Tabungan (Livin'), PDF berpassword - isi kolom Password |
| BSI | Bank Syariah Indonesia | Account Statement |
| CIMB | CIMB Niaga | BizChannel Account Statement |
| MAYBANK | Maybank Indonesia | Cash Management Account Report |
| BPD_DIY | BPD DIY | Rekening Koran |
| BPR_NATASHA | BPR Natasha Bintang Anugrah | Rekening Koran Tabungan |

## Arti status verifikasi

| Status | Arti |
|---|---|
| OK | Saldo berjalan, saldo awal/akhir, dan total debit/kredit cocok |
| Perlu dicek (WARN) | Contoh: saldo awal tidak tercantum di PDF sehingga dihitung, tanggal di luar periode, bulan hilang, atau file duplikat |
| Gagal (ERROR) | Angka tidak cocok, atau PDF tidak bisa dibaca |

File hasil tetap dibuat walaupun ada peringatan. Rinciannya ada di **Log proses**.

## Jika ada kendala

| Masalah | Solusi |
|---|---|
| "BSP.exe sedang terbuka" saat instalasi | Tutup aplikasi BSP, lalu jalankan `Install_BSP.bat` lagi. |
| "R atau engine BSP belum terpasang" di aplikasi | Jalankan `Install_BSP.bat` di folder program. |
| Paket R gagal dipasang | Periksa internet atau proxy kantor, atau pakai cara PC tanpa internet di atas. |
| Kebijakan kantor memblokir exe buatan sendiri | Hubungi IT; BSP.exe dibuat oleh skrip di folder program. |
| Shortcut desktop tidak berfungsi | Folder program dipindah. Jalankan `Install_BSP.bat` lagi dari lokasi baru. |
| PDF hasil scan (gambar) | Belum didukung dan akan ditandai Gagal. |

Log instalasi: `setup\install_log.txt`. Log engine: `%TEMP%\BSP_work\engine_log.txt`.

## Isi folder

| File | Fungsi |
|---|---|
| `Install_BSP.bat` | Instalasi 1 klik |
| `setup\install.ps1`, `setup\build_exe.ps1` | Langkah instalasi dan pembuatan BSP.exe |
| `app\BSP.cs`, `app\BSP.ico`, `app\logo.png` | Sumber aplikasi, ikon, dan logo |
| `BSP_modular_engine\` | Engine R: `Core\` (proses) dan `Banks\` (satu file per bank) |
