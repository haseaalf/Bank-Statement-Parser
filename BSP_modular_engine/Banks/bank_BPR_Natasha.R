#' PT BPR Natasha Bintang Anugrah - Laporan Rekening Koran Tabungan
#'
#' Row layout : No | Tgl Trans (dd/mm/yyyy) | Uraian | sandi | Debet | Kredit | Saldo | User ID
#'              Amounts without decimals, dot as thousand separator (32.989.696)
#' Summary    : "Saldo Awal" row and "SALDO AKHIR" row (total debit, total kredit, saldo)

natasha_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("BPR NATASHA", "LAPORAN REKENING KORAN", "Tgl Trans", "sandi", "SALDO AKHIR"),
                required = "NATASHA")
}

natasha_parse <- function(doc) {
  L <- doc$lines
  txt <- doc_all_text(doc)

  name_line <- which(grepl("^Nama Nasabah", L$text))[1]
  nw <- line_words(doc, L$line_id[name_line])
  pd <- grab(txt, "Periode\\s*:?\\s*(\\d{1,2} \\w+ \\d{4}\\s*-\\s*\\d{1,2} \\w+ \\d{4})")
  if (is.na(pd)) pd <- paste(nw$text[nw$x > 420], collapse = " ")   # value printed on the name line
  pdp <- strsplit(sub("^:\\s*", "", pd), "\\s*-\\s*")[[1]]
  meta <- list(account_no = grab(txt, "No\\. Rekening\\s*:?\\s*(\\d+)"),
               account_name = paste(nw$text[nw$x > 105 & nw$x < 340], collapse = " "),
               period_start = parse_text_date(pdp[1]),
               period_end = parse_text_date(if (length(pdp) > 1) pdp[2] else NA))

  hdr_re <- "^No\\. Tgl Trans Uraian"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  sandi_x <- th$x[th$text == "sandi"]
  debet_x <- th$x[th$text == "Debet"] - 15
  user_x <- th$x[th$text == "User"] - 5

  body <- table_lines(doc, header = hdr_re, stop = "^SALDO BLOKIR|^Halaman")
  stated <- list()
  rows <- list()
  for (id in body$line_id) {
    w <- line_words(doc, id)
    w <- w[w$x < user_x, ]
    line_txt <- paste(w$text, collapse = " ")
    amt <- w[w$x >= debet_x & is_amount(w$text, "id"), ]
    if (grepl("^Saldo Awal", line_txt)) {
      stated$opening <- parse_amount(tail(amt$text, 1), "id"); next
    }
    if (grepl("^SALDO AKHIR", line_txt)) {
      v <- parse_amount(amt$text, "id")
      if (length(v) == 3) stated <- c(stated, list(total_debit = v[1], total_credit = v[2], closing = v[3]))
      next
    }
    if (!(grepl("^\\d+$", w$text[1]) && grepl("^\\d\\d/\\d\\d/\\d{4}$", w$text[2]))) {
      # wrapped Uraian
      if (length(rows) > 0 && all(w$x > th$x[th$text == "Uraian"] - 80 & w$x < sandi_x)) {
        rows[[length(rows)]]$desc <- c(rows[[length(rows)]]$desc, line_txt)
      }
      next
    }
    v <- parse_amount(tail(amt$text, 3), "id")
    is_sandi <- grepl("^\\d{2}$", w$text) & w$x >= sandi_x - 10
    desc_w <- w[seq_len(nrow(w)) > 2 & w$x < debet_x & !is_sandi, ]
    desc <- desc_w$text
    # the sandi code is sometimes glued to the last word ("JAYA08")
    n <- length(desc)
    if (n > 0 && desc_w$x2[n] > sandi_x && grepl("[A-Za-z]\\d{2}$", desc[n])) desc[n] <- sub("\\d{2}$", "", desc[n])
    rows[[length(rows) + 1]] <- list(date = parse_num_date(w$text[2]), desc = paste(desc, collapse = " "),
                                     debit = v[1], credit = v[2], balance = v[3], page = w$page[1])
  }

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "BPR_NATASHA", name = "BPR Natasha Bintang Anugrah", variant = "Rekening Koran Tabungan",
              description = "Laporan rekening koran tabungan BPR Natasha: No, Tgl Trans, Uraian, sandi, Debet, Kredit, Saldo",
              detect = natasha_detect, parse = natasha_parse)
