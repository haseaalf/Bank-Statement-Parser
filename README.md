# Bank Statement Parser (BSP)

Excel macro (VBA) + engine R yang mengubah PDF rekening koran berbagai bank menjadi Excel
dengan format standar, lalu memverifikasi kewajaran saldonya.

Format output selalu sama dengan sampel `input output sample/*.xlsx`: satu sheet bernama kode bank,
kolom `Transaction Date (DD/MM)` | `Description` | `Debit` | `Credit` | `Balance`.
Format ini juga sama dengan input *Append Monthly Statement* di BSA.

Arsitekturnya mengikuti BSA: workbook memanggil `Rscript.exe` dengan file job di folder kerja,
engine membaca semua script di foldernya, lalu hasilnya dibaca kembali oleh VBA.

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

## Struktur

```
BSP_v1.xlsm                      UI Excel (dibangun dari excel/src dengan excel/build_xlsm.ps1)
BSP_modular_engine/
  BSP_engine.R                   entry: Rscript BSP_engine.R <folder kerja> <folder engine> <list_banks|detect|parse>
  Core/                          baca PDF, util angka/tanggal, registry, deteksi, verifikasi, konsolidasi, writer xlsx
  Banks/                         satu file per bank (plugin) -> bank_<KODE>.R
excel/src/*.bas                  modul VBA (Mdl_BSP_Engine, Mdl_BSP_UI, Mdl_BSP_Setup)
tests/run_tests.R                tes golden BRI + semua sampel + konsolidasi + engine end-to-end
tools/dump_lines.R               alat bantu membuat parser: tampilkan baris & posisi kata sebuah PDF
docs/bank_template.R             template parser bank baru
```

## Instalasi 1 klik (untuk pengguna)

1. Salin atau ekstrak seluruh folder program ke PC.
2. Double-click **`Install_BSP.bat`**. Hak administrator tidak diperlukan. Skrip ini akan:
   - memeriksa Windows dan Microsoft Excel;
   - mencari R (versi 4.1 ke atas). Jika belum ada, R **4.2.2** (versi yang sudah teruji) diunduh dari
     CRAN dan dipasang khusus untuk user tersebut di `%LOCALAPPDATA%\Programs\R`;
   - memasang paket `pdftools`, `openxlsx`, dan `zip` ke `BSP_modular_engine\r-library\<versi R>`;
   - melepas blokir "file dari internet" pada file program, supaya macro `BSP_v1.xlsm` tidak diblokir;
   - menulis `bsp_config.txt` (lokasi R dan engine) dan membuat shortcut **Bank Statement Parser** di
     desktop;
   - menjalankan engine sekali sebagai uji coba.
3. Buka shortcut tersebut, klik **Enable Content** bila diminta, lalu **Perbarui daftar bank**.

Saat dibuka, workbook memperbaiki sendiri lokasi R dan engine yang tidak ditemukan, misalnya karena
folder dipindah atau disalin dari PC lain. Log instalasi ada di `setup\install_log.txt`.

**PC tanpa internet:** di PC yang punya internet, jalankan `Install_BSP.bat -PrepareOffline`. Perintah
ini mengunduh installer R dan paket R ke `setup\offline\`. Salin seluruh folder program ke PC tujuan,
lalu jalankan `Install_BSP.bat` di sana; skrip otomatis memakai file offline tersebut.

Opsi lain: `-RVersion 4.x.y` (versi R yang dipasang bila R belum ada) dan `-NoShortcut`.

## Persiapan untuk developer

1. **R** (sudah dites dengan R 4.2.2). Jalankan `Install_BSP.bat`, atau biarkan engine memasang
   paketnya sendiri saat pertama kali berjalan.
2. **Bangun `BSP_v1.xlsm`.**
   - Di Excel, aktifkan *File > Options > Trust Center > Trust Center Settings > Macro Settings >
     Trust access to the VBA project object model*.
   - Jalankan:
     ```
     powershell -ExecutionPolicy Bypass -File excel\build_xlsm.ps1 -Test
     ```
   - Pengaturan tersebut boleh dimatikan lagi setelah build.
   - Cara manual: buat workbook baru, buka Alt+F11, *File > Import File* untuk ketiga file
     `excel/src/*.bas`, jalankan macro `BSP_Setup_UI`, lalu simpan sebagai `.xlsm`.

## Cara pakai

1. **Pengaturan engine R.**
   - Folder R (berisi `Rscript.exe`) dan folder engine terisi otomatis jika ditemukan. Jika tidak,
     pilih lewat tombol *Browse*.
   - Klik **Perbarui daftar bank** untuk mengisi pilihan bank.
2. **Output.**
   - Pilih folder tujuan, atau pilih *Sama dengan folder PDF*.
   - Mode output: *Per file* (1 PDF menjadi 1 Excel) atau *Gabung per rekening*.
   - Pilih juga perlakuan jika file output sudah ada.
   - **Nama file output**: *Sama dengan nama PDF* atau *Standar* (`<BANK>_<NoRek>_<yyyymmdd>-<yyyymmdd>.xlsx`).
     Mode *Gabung per rekening* selalu memakai nama standar.
3. **File.**
   - Klik **Pilih File PDF...** (bisa banyak file sekaligus), lalu **Deteksi Bank**.
   - Kolom *Bank* = `AUTO` berarti deteksi otomatis. Ubah ke kode bank tertentu jika deteksi salah
     atau ragu.
   - Isi *Password* jika PDF terkunci.
4. **PROSES.**
   - Hasil per file tampil di sheet **Hasil**, lengkap dengan link ke file output.
   - Temuan verifikasi tampil di sheet **Verifikasi**.
   - **Verifikasi saja** menjalankan pengecekan tanpa membuat file output.

### Arti verifikasi

| Tingkat | Arti |
|---|---|
| ERROR | Angka tidak cocok: saldo berjalan, saldo awal/akhir, total debit/kredit, jumlah transaksi, atau saldo antar periode tidak bersambung |
| WARN | Perlu dicek: saldo awal dihitung (tidak tercantum di PDF), tanggal di luar periode, deteksi ragu, periode hilang/tumpang tindih, file duplikat |
| INFO | Catatan, misalnya saldo akhir tidak tercantum |

Output tetap dibuat walaupun ada ERROR/WARN. Saldo yang tidak dicetak bank (BCA hanya mencetak saldo
akhir hari; baris terakhir BNI) dihitung lalu dicocokkan dengan saldo yang tercetak.

### Rekap mutasi

Pilih **Buat rekap mutasi = Ya** sebelum **PROSES**. Semua PDF dalam satu proses (boleh banyak bank, rekening,
dan nasabah sekaligus, dalam urutan apa pun) diringkas ke satu file `<yyyymmdd>_<hhmm>_rekap_mutasi.xlsx`:

- **REKAP**: Bank, Nomor Rekening (teks, nol di depan tetap ada), Nama Nasabah, Debit, Kredit, Saldo, %.
  - Debit dan Kredit adalah total semua bulan.
  - Saldo adalah rata-rata saldo akhir bulan.
  - % adalah porsi saldo terhadap jumlah semua rekening.
  - Di bawahnya ada baris Jumlah, *Penjualan L/R Periode ...* (isi manual), dan *% Transaksi Bank*
    (= total kredit / penjualan, otomatis terhitung setelah penjualan diisi).
- **Satu sheet per rekening** (nama sheet = nomor rekening), berisi BULAN | DEBET | KREDIT | SALDO AKHIR BULAN,
  diikuti baris Total dan Ratas. Rentang bulannya sama untuk semua rekening.
- **Verifikasi per bulan**: saldo akhir bulan lalu + kredit - debit harus sama dengan saldo akhir bulan ini.
  Masalah ditulis sebagai `PERINGATAN` di kolom E di samping baris bulannya. Contohnya selisih saldo,
  bulan yang tidak di-upload, bulan yang tidak bersambung, file duplikat, dan periode tumpang tindih.
  Ringkasannya juga muncul di kolom H sheet REKAP, serta di sheet Hasil/Verifikasi BSP.

### Konsolidasi

- File digabung per **bank + nomor rekening + mata uang**, lalu diurutkan menurut periode.
- Saldo akhir periode sebelumnya harus sama dengan saldo awal periode berikutnya.
- File duplikat tidak ikut digabung.
- Nama output: `<BANK>_<NoRek>_<yyyymmdd>-<yyyymmdd>.xlsx`.

## Menambah bank baru

1. Salin `docs/bank_template.R` ke `BSP_modular_engine/Banks/bank_<KODE>.R`.
2. Lihat struktur PDF-nya dengan:
   ```
   Rscript tools/dump_lines.R "file.pdf"
   ```
3. Isi tiga bagian:
   - `detect()`: frasa khas bank tersebut.
   - `parse()`: baca meta, ringkasan, dan transaksi menjadi `new_statement()`.
   - `register_bank()`.
4. Tambahkan pola nama file ke `tests/expected_banks.csv`, lalu jalankan `Rscript tests/run_tests.R <KODE>`.
5. Klik **Perbarui daftar bank** di Excel. VBA tidak perlu diubah.

Fungsi bantu yang tersedia antara lain `read_statement_pdf`, `line_words`, `table_lines`,
`nearest_anchor`, `table_row_rules`, `parse_amount` (style `"us"`/`"id"`), `parse_num_date`,
`parse_text_date`, `date_from_day_month`, `fix_order_by_chain`, `keyword_score`, dan `rows_to_txns`.

## Tes

```
"C:\Program Files\R\R-4.2.2\bin\Rscript.exe" tests\run_tests.R
```

PDF sampel bersifat rahasia dan tidak ikut di git (lihat `.gitignore`). Bagian tes yang butuh sampel
dilewati jika foldernya tidak ada.

## Catatan

- Sampel output BRI lama berasal dari konverter OCR, jadi beberapa deskripsinya salah ketik,
  misalnya `DEWI WUI` padahal di PDF tertulis `DEWI WIJI`, atau `SE11YOWATI` untuk `SETIYOWATI`.
  BSP mengikuti teks asli PDF. Tanggal, debit, kredit, dan saldo identik dengan sampel.
- PDF hasil scan (tanpa lapisan teks) belum didukung dan akan ditandai ERROR.
- Sampel bank selain BRI/BCA/BNI hanya 1 halaman. PDF multi-halaman dari bank tersebut perlu diuji
  saat tersedia.
