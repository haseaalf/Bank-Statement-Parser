#' BNI - Laporan Mutasi Rekening (BNI mobile banking, PDF with password)
#'
#' Row layout : 3 lines per transaction
#'                "01 Oct 2026"  + category (Transfer / Pembayaran Qris / Biaya ...)
#'                signed amount (+ credit / - debit) + balance       <- anchor line
#'                "08:12:19 WIB" + details
#'              Amounts with comma thousands, usually without decimals (2,886,838).
#'              Table opens with a "Saldo Awal" row and closes with a "Saldo Akhir" row.
#' Summary    : page 1 - Saldo Awal, Total Pemasukan, Total Pengeluaran, Saldo Akhir

bni_mutasi_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("Laporan Mutasi Rekening", "Total Pemasukan", "Total Pengeluaran", "Rincian Transaksi",
                  "Nominal \\(IDR\\)", "Bank Negara Indonesia"),
                required = "Total Pemasukan")
}

bni_mutasi_parse <- function(doc) {
  L <- doc$lines
  txt <- gsub("\\s+", " ", doc_all_text(doc))

  # ---- Meta: name + "<product> - <account>" on one line, period "1 - 31 Oktober 2026"
  hdr_re <- "^Tanggal & Waktu Rincian Transaksi"
  first_hdr <- which(grepl(hdr_re, L$text))[1]
  acct_line <- which(L$page == 1 & grepl("- \\d{6,}$", L$text) & seq_len(nrow(L)) < first_hdr)[1]
  aw <- line_words(doc, L$line_id[acct_line])
  split_x <- min(aw$x[aw$x > 300])
  per <- regmatches(txt, regexec("Periode:\\s*(\\d{1,2})(?:\\s+(\\w+))?(?:\\s+(\\d{4}))?\\s*-\\s*(\\d{1,2})\\s+(\\w+)\\s+(\\d{4})",
                                 txt, perl = TRUE))[[1]]
  ps <- pe <- as.Date(NA)
  if (length(per) == 7) {
    pe <- parse_text_date(paste(per[5], per[6], per[7]))
    ps <- parse_text_date(paste(per[2], if (nzchar(per[3])) per[3] else per[6], if (nzchar(per[4])) per[4] else per[7]))
  }
  meta <- list(account_no = grab(L$text[acct_line], "(\\d{6,})$"),
               account_name = paste(aw$text[aw$x < split_x], collapse = " "),
               currency = grab(txt, "Mata Uang:\\s*([A-Z]{3})"),
               period_start = ps, period_end = pe)

  # ---- Summary: label line "Saldo Awal Total Pemasukan Total Pengeluaran Saldo Akhir", values below
  stated <- list()
  si <- which(grepl("^Saldo Awal Total Pemasukan", L$text))[1]
  if (!is.na(si)) {
    v <- strsplit(L$text[si + 1], " ")[[1]]
    v <- parse_amount(sub("^[+-]", "", v[grepl("^[+-]?[0-9][0-9,.]*$", v)]))
    if (length(v) == 4) stated <- list(opening = v[1], total_credit = v[2], total_debit = v[3], closing = v[4])
  }

  # ---- Transactions
  th <- line_words(doc, L$line_id[first_hdr])
  det_x <- th$x[th$text == "Rincian"] - 90        # details start left of the centred header
  amt_x <- th$x[th$text == "Nominal"] - 40
  bal_x <- th$x[th$text == "Saldo"] - 45
  body <- table_lines(doc, header = hdr_re, stop = "^Saldo Akhir|^PT Bank Negara|^Informasi Lainnya")
  body <- body[!grepl("^Saldo Awal", body$text), ]
  signed <- "^[+-][0-9][0-9,]*(\\.[0-9]+)?$"
  body$anchor <- vapply(body$line_id, function(id) {
    w <- line_words(doc, id)
    any(grepl(signed, w$text[w$x >= amt_x & w$x < bal_x])) && any(w$x >= bal_x & is_amount(w$text))
  }, logical(1))

  rows <- list()
  for (p in unique(body$page)) {
    pb <- body[body$page == p, ]
    a <- pb[pb$anchor, ]
    if (nrow(a) == 0) {
      if (length(rows) > 0) rows[[length(rows)]]$lines <- rbind(rows[[length(rows)]]$lines, pb)
      next
    }
    owner <- nearest_anchor(pb, a)
    # lines far above the first anchor continue the last transaction of the previous page
    carry <- pb$y < a$y[1] - 15
    if (any(carry) && length(rows) > 0) rows[[length(rows)]]$lines <- rbind(rows[[length(rows)]]$lines, pb[carry, ])
    for (k in seq_len(nrow(a))) rows[[length(rows) + 1]] <- list(anchor = a$line_id[k], page = p, lines = pb[owner == k & !carry, ])
  }

  txns <- lapply(rows, function(r) {
    blk <- r$lines[order(r$lines$page, r$lines$y), ]
    date_txt <- character(); details <- character()
    for (id in blk$line_id) {
      w <- line_words(doc, id)
      date_txt <- c(date_txt, w$text[w$x < det_x])
      d <- paste(w$text[w$x >= det_x & w$x < amt_x], collapse = " ")
      if (nchar(d)) details <- c(details, d)
    }
    aw <- line_words(doc, r$anchor)
    amt <- aw$text[aw$x >= amt_x & aw$x < bal_x & grepl(signed, aw$text)][1]
    v <- parse_amount(sub("^[+-]", "", amt))
    list(date = parse_text_date(paste(date_txt, collapse = " ")), desc = details,
         debit = if (startsWith(amt, "-")) v else 0, credit = if (startsWith(amt, "+")) v else 0,
         balance = parse_amount(aw$text[aw$x >= bal_x & is_amount(aw$text)][1]), page = r$page)
  })

  new_statement(rows_to_txns(txns), meta, stated)
}

register_bank(id = "BNI", name = "Bank Negara Indonesia (BNI)", variant = "Laporan Mutasi Rekening (mobile)",
              description = "Laporan mutasi BNI mobile banking: Tanggal & Waktu, Rincian, Nominal (+/-), Saldo; PDF berpassword",
              detect = bni_mutasi_detect, parse = bni_mutasi_parse)
