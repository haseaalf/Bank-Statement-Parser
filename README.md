# Bank Statement Parser (BSP)

Excel macro + engine R yang mengubah PDF rekening koran berbagai bank menjadi Excel dengan format
standar, memverifikasi kewajaran saldonya, dan bisa membuat rekap mutasi bulanan.

Format output: satu sheet bernama kode bank, kolom `Transaction Date (DD/MM)` | `Description` |
`Debit` | `Credit` | `Balance`. Format ini sama dengan input *Append Monthly Statement* di BSA.

> Branch `release` hanya berisi file yang dibutuhkan untuk **memakai** program.
> Kode sumber VBA, tes, dan alat bantu developer ada di branch `feature/bsp-v1` / `main`.

## Isi folder

```
Install_BSP.bat          instalasi 1 klik (double-click)
BSP_v1.xlsm              program (Excel)
BSP_modular_engine/      engine R: Core/ (proses) dan Banks/ (satu file per bank)
setup/install.ps1        dipakai oleh Install_BSP.bat
```

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

## Instalasi (sekali per PC)

1. Unduh branch ini (*Code > Download ZIP*) lalu ekstrak seluruh isinya ke satu folder.
2. Double-click **`Install_BSP.bat`**. Hak administrator tidak diperlukan. Skrip ini akan:
   - memeriksa Windows dan Microsoft Excel;
   - mencari R (versi 4.1 ke atas). Jika belum ada, R **4.2.2** diunduh dari CRAN dan dipasang khusus
     untuk user tersebut;
   - memasang paket R yang dibutuhkan (`pdftools`, `openxlsx`, `zip`);
   - melepas blokir "file dari internet" supaya macro tidak diblokir Excel;
   - membuat shortcut **Bank Statement Parser** di desktop, lalu menguji engine sekali.
3. Buka shortcut tersebut, klik **Enable Content** bila diminta, lalu klik **Perbarui daftar bank**.

**PC tanpa internet:** di PC yang online, jalankan `Install_BSP.bat -PrepareOffline`. Perintah ini
mengunduh installer R dan paket R ke `setup\offline\`. Salin seluruh folder ke PC tujuan, lalu
jalankan `Install_BSP.bat` di sana.

Jika instalasi gagal, lihat `setup\install_log.txt`.

## Cara pakai

1. **Output.**
   - Pilih folder tujuan, atau pilih *Sama dengan folder PDF*.
   - Mode output: *Per file* (1 PDF menjadi 1 Excel) atau *Gabung per rekening*.
   - Pilih perlakuan jika file output sudah ada.
   - **Nama file output**: *Sama dengan nama PDF* atau *Standar* (`<BANK>_<NoRek>_<yyyymmdd>-<yyyymmdd>.xlsx`).
     Mode gabung selalu memakai nama standar.
   - **Buat rekap mutasi**: *Ya* untuk membuat rekap bulanan semua file (lihat di bawah).
2. **File.**
   - Klik **Pilih File PDF...** (bisa banyak file, banyak bank, dan banyak nasabah sekaligus), lalu
     **Deteksi Bank**.
   - Kolom *Bank* = `AUTO` berarti deteksi otomatis. Ubah ke kode bank tertentu jika status
     "Tidak dikenali" atau "Ragu".
   - Isi *Password* jika PDF terkunci.
3. **PROSES.**
   - Hasil per file tampil di sheet **Hasil**, lengkap dengan link ke file output.
   - Temuan verifikasi tampil di sheet **Verifikasi**.
   - **Verifikasi saja** menjalankan pengecekan tanpa membuat file output.

### Arti verifikasi

| Tingkat | Arti |
|---|---|
| ERROR | Angka tidak cocok: saldo berjalan, saldo awal/akhir, total debit/kredit, jumlah transaksi, atau saldo antar periode tidak bersambung |
| WARN | Perlu dicek: saldo awal dihitung (tidak tercantum di PDF), tanggal di luar periode, deteksi ragu, periode hilang/tumpang tindih, file duplikat |
| INFO | Catatan, misalnya tidak ada transaksi pada periode itu |

Output tetap dibuat walaupun ada ERROR/WARN.

Kolom **Skor** di daftar file menunjukkan keyakinan deteksi bank (0–1): 1 = semua ciri bank cocok,
di bawah 0,5 = tidak dikenali (pilih bank secara manual).

### Rekap mutasi

Semua PDF dalam satu proses diringkas ke satu file `<yyyymmdd>_<hhmm>_rekap_mutasi.xlsx`:

- **REKAP**: Bank, Nomor Rekening, Nama Nasabah, Debit, Kredit (total semua bulan), Saldo (rata-rata
  saldo akhir bulan), dan % (porsi saldo).
  - Baris *Penjualan L/R Periode ...* diisi manual.
  - *% Transaksi Bank* (= total kredit / penjualan) langsung terhitung setelah penjualan diisi.
- **Satu sheet per rekening**: BULAN | DEBET | KREDIT | SALDO AKHIR BULAN, ditambah baris Total dan Ratas.
- **Cek bulanan**: saldo akhir bulan lalu + kredit - debit harus sama dengan saldo akhir bulan ini.
  Masalah ditulis sebagai `PERINGATAN` di kolom E di samping bulannya, misalnya bulan yang tidak
  di-upload, bulan tidak bersambung, atau file duplikat.

### Gabung per rekening

- File dari bank + nomor rekening + mata uang yang sama digabung dan diurutkan menurut periode.
- Saldo akhir periode sebelumnya harus sama dengan saldo awal periode berikutnya.
- File duplikat tidak ikut digabung.

## Catatan

- PDF hasil scan (gambar, tanpa teks) belum didukung dan akan ditandai ERROR.
- Deskripsi transaksi mengikuti teks asli di PDF.
- File rekening koran berisi data rahasia nasabah: jangan menyimpannya di folder program yang dibagikan.
