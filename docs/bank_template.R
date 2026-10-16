#' <NAMA BANK> - <nama laporan / jenis statement>
#'
#' Salin file ini ke BSP_modular_engine/Banks/bank_<KODE>.R lalu sesuaikan.
#' Engine otomatis membaca semua file di folder Banks/, jadi tidak ada kode lain
#' (R maupun VBA) yang perlu diubah. Setelah itu klik "Perbarui daftar bank" di Excel.
#'
#' Row layout : <kolom-kolom tabel transaksi>
#' Summary    : <angka ringkasan yang tercantum: saldo awal/akhir, total, jumlah transaksi>
#'
#' Alat bantu:
#'   Rscript tools/dump_lines.R "<file.pdf>"   -> lihat baris + posisi x setiap kata
#'   Rscript tests/run_tests.R <filter-nama>   -> jalankan tes untuk file tertentu

xxx_detect <- function(doc) {
  # Skor 0-1. Pakai frasa yang khas untuk statement bank ini.
  keyword_score(doc_all_text(doc),
                c("FRASA KHAS 1", "FRASA KHAS 2", "JUDUL KOLOM"),
                required = "FRASA WAJIB")
}

xxx_parse <- function(doc) {
  L <- doc$lines
  txt <- doc_all_text(doc)

  # ---- Meta: no. rekening, nama, mata uang, periode -------------------------
  meta <- list(account_no = grab(txt, "No\\. Rekening\\s*:\\s*(\\d+)"),
               account_name = NA_character_,
               currency = "IDR",
               period_start = parse_num_date(grab(txt, "Periode\\s*:\\s*(\\S+)")),
               period_end = parse_num_date(grab(txt, "s/d\\s*(\\S+)")))

  # ---- Ringkasan dari PDF (isi NA bila tidak ada) ---------------------------
  stated <- list(opening = parse_amount(grab(txt, "Saldo Awal\\s*:\\s*([0-9.,]+)")),
                 closing = parse_amount(grab(txt, "Saldo Akhir\\s*:\\s*([0-9.,]+)")))

  # ---- Transaksi -------------------------------------------------------------
  # Posisi kolom diambil dari baris judul tabel supaya tahan terhadap geseran kecil
  hdr_re <- "^Tanggal Keterangan"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  desc_x <- th$x[th$text == "Keterangan"] - 5
  debit_x <- th$x[th$text == "Debet"] - 30

  body <- table_lines(doc, header = hdr_re, stop = "^Saldo Akhir")
  rows <- list(); cur <- NULL
  for (id in body$line_id) {
    w <- line_words(doc, id)
    if (grepl("^\\d\\d/\\d\\d/\\d{4}$", w$text[1])) {        # baris baru = diawali tanggal
      if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
      amt <- tail(which(is_amount(w$text) & w$x >= debit_x), 3)  # debit, kredit, saldo
      v <- parse_amount(w$text[amt])                              # style = "id" untuk 1.234,56
      cur <- list(date = parse_num_date(w$text[1]),
                  desc = paste(w$text[w$x >= desc_x & w$x < debit_x], collapse = " "),
                  debit = v[1], credit = v[2], balance = v[3], page = w$page[1])
    } else if (!is.null(cur)) {                                   # lanjutan deskripsi
      cur$desc <- c(cur$desc, paste(w$text[w$x >= desc_x & w$x < debit_x], collapse = " "))
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur

  # Bila bank mencetak transaksi terbaru di atas, panggil fix_order_by_chain()
  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "XXX",                         # kode singkat = nama sheet output
              name = "Nama Bank Lengkap",
              variant = "Nama Format",            # satu bank boleh punya beberapa format
              description = "Kolom: Tanggal, Keterangan, Debet, Kredit, Saldo",
              detect = xxx_detect, parse = xxx_parse)
