#' BCA - e-Statement Rekening Giro / Tahapan (Laporan Mutasi Rekening)
#'
#' Row layout : dd/mm (no year) | KETERANGAN (type + details, many lines) | CBG |
#'              MUTASI (+ "DB" for debit) | SALDO (printed only on some rows)
#' Year       : from "PERIODE : FEBRUARI 2026"
#' Summary    : last page - SALDO AWAL, MUTASI CR (amount + count),
#'              MUTASI DB (amount + count), SALDO AKHIR

bca_detect <- function(doc) {
  t <- gsub(" ", "", doc_all_text(doc))      # header fonts are letter-spaced
  keyword_score(t, c("REKENINGGIRO|REKENINGTAHAPAN|REKENINGKORAN", "KETERANGAN", "MUTASI",
                     "SALDOAWAL", "BCAberhak", "Bersambungkehalamanberikut"),
                required = "BCA")
}

#' Join words whose boxes almost touch (letter-spaced header text)
merge_tight_words <- function(w, gap = 1.2) {
  if (nrow(w) == 0) return(character())
  out <- w$text[1]
  for (i in seq_len(nrow(w))[-1]) {
    if (w$x[i] - w$x2[i - 1] <= gap) out[length(out)] <- paste0(out[length(out)], w$text[i])
    else out <- c(out, w$text[i])
  }
  out
}

bca_parse <- function(doc) {
  L <- doc$lines
  compact <- gsub(" ", "", L$text)

  # ---- Meta (first page header)
  rek <- which(grepl("NO\\.REKENING:", compact))[1]
  rw <- line_words(doc, L$line_id[rek])
  per <- grab(compact[which(grepl("PERIODE:", compact))[1]], "PERIODE:([A-Za-z]+)(\\d{4})", 1)
  year <- grab(compact[which(grepl("PERIODE:", compact))[1]], "PERIODE:([A-Za-z]+)(\\d{4})", 2)
  m <- month_number(per)
  ps <- if (!is.na(m)) make_date(year, m, 1) else as.Date(NA)
  pe <- if (!is.na(m)) seq(ps, by = "month", length.out = 2)[2] - 1 else as.Date(NA)
  meta <- list(account_no = grab(compact[rek], "NO\\.REKENING:(\\d+)"),
               account_name = paste(merge_tight_words(rw[rw$x < 300, ]), collapse = " "),
               currency = grab(compact[which(grepl("MATAUANG:", compact))[1]], "MATAUANG:([A-Z]{3})"),
               period_start = ps, period_end = pe)

  # ---- Summary (last page)
  stated <- list()
  sv <- function(label) {
    i <- which(grepl(paste0("^", label, ":"), compact))
    if (length(i) == 0) return(c(NA, NA))
    w <- line_words(doc, L$line_id[tail(i, 1)])
    a <- w$text[is_amount(w$text)]
    c(parse_amount(a[1]), if (length(a) > 1) as.numeric(gsub(",", "", a[2])) else NA)
  }
  stated$opening <- sv("SALDOAWAL")[1]
  cr <- sv("MUTASICR"); db <- sv("MUTASIDB")
  stated$total_credit <- cr[1]; stated$n_credit <- as.integer(cr[2])
  stated$total_debit <- db[1]; stated$n_debit <- as.integer(db[2])
  stated$closing <- sv("SALDOAKHIR")[1]

  # ---- Transactions
  hdr_re <- "^TANGGALKETERANGAN"
  hi <- which(grepl(hdr_re, compact))[1]
  th <- line_words(doc, L$line_id[hi])
  ket_x <- min(th$x[th$x > 100]) - 80            # first KETERANGAN column (type)
  det_x <- th$x[grepl("^C$|^CBG", th$text)][1] - 120  # second KETERANGAN column (details)
  cbg_x <- th$x[grepl("^C$|^CBG", th$text)][1] - 5
  mut_x <- cbg_x + 40
  saldo_x <- th$x[grepl("^SALD", th$text)][1] - 15

  body <- table_lines(doc, header = hdr_re, compact_header = TRUE,
                      stop = "Bersambung ke halaman|^SALDO AWAL :")
  rows <- list(); cur <- NULL
  for (id in body$line_id) {
    w <- line_words(doc, id)
    if (grepl("^\\d\\d/\\d\\d$", w$text[1])) {
      if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
      cur <- NULL
      typ <- paste(w$text[w$x >= ket_x & w$x < det_x], collapse = " ")
      if (grepl("^SALDO AWAL", typ)) {
        if (is.na(stated$opening)) stated$opening <- parse_amount(tail(w$text, 1))
        next
      }
      mw <- w[w$x >= mut_x & w$x < saldo_x, ]
      amt_i <- which(is_amount(mw$text))
      amount <- parse_amount(mw$text[amt_i[1]])
      is_db <- any(mw$text == "DB") || grepl("DB$", mw$text[amt_i[1]])
      sw <- w[w$x >= saldo_x & is_amount(w$text), ]
      det <- paste(w$text[w$x >= det_x & w$x < cbg_x], collapse = " ")
      cur <- list(date = date_from_day_month(w$text[1], ps, pe),
                  desc = paste(c(typ, det)[nchar(c(typ, det)) > 0], collapse = " "),
                  debit = if (is_db) amount else 0, credit = if (is_db) 0 else amount,
                  balance = if (nrow(sw) > 0) parse_amount(sw$text[1]) else NA, page = w$page[1])
    } else if (!is.null(cur)) {
      txt <- paste(w$text[w$x >= ket_x & w$x < cbg_x], collapse = " ")
      if (nchar(txt)) cur$desc <- c(cur$desc, txt)
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "BCA", name = "Bank Central Asia (BCA)", variant = "e-Statement Rekening Giro",
              description = "e-Statement BCA: TANGGAL (DD/MM), KETERANGAN, CBG, MUTASI (DB), SALDO",
              detect = bca_detect, parse = bca_parse)
