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

## Persiapan (sekali)

1. **R** (sudah dites dengan R 4.2.2). Paket `pdftools` dan `openxlsx` terpasang otomatis saat engine
   pertama kali berjalan. Jika internet kantor memblokir CRAN, taruh paketnya di
   `BSP_modular_engine/r-library/`.
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
