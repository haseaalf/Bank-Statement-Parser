#' BPD DIY (Bank Pembangunan Daerah DI Yogyakarta) - Rekening Koran
#'
#' Row layout : Tgl Val | Tgl Tx | Keterangan | Tx | No Arsip | Debet | Kredit | Saldo + K/D
#'              Indonesian number format (9.253.601,67). A row may have no date
#'              (same date as the row above).
#' Summary    : "Saldo Awal :" row and "J u m l a h :" row (total debit, total credit, saldo akhir)

bpd_diy_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("R E K E N I N G K O R A N", "Tgl Val", "No Arsip", "KTX:", "BPD DIY", "Saldo Awal"),
                required = "Tgl Val")
}

#' "9.253.601,67" + "K" -> 9253601.67 ; "D" (debit balance) -> negative
bpd_balance <- function(value, marker) {
  v <- parse_amount(value, style = "id")
  if (identical(marker, "D")) -v else v
}

bpd_diy_parse <- function(doc) {
  L <- doc$lines
  txt <- doc_all_text(doc)

  rek_line <- which(grepl("NO\\. REK", L$text))[1]
  name_words <- line_words(doc, L$line_id[rek_line - 1])
  meta <- list(account_no = grab(txt, "NO\\. REK\\s*:\\s*([0-9.]+)"),
               account_name = paste(name_words$text[name_words$x < 300], collapse = " "),
               period_start = parse_num_date(grab(txt, "PERIODE\\s*:\\s*(\\S+)\\s+S/D")),
               period_end = parse_num_date(grab(txt, "S/D\\s+(\\S+)")))

  hdr_re <- "^Tgl Val Tgl Tx Keterangan"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  tx_x <- th$x[th$text == "Tgl"][2] - 5
  ket_x <- th$x[th$text == "Keterangan"] - 45
  code_x <- th$x[th$text == "Tx"][2] - 12
  arsip_x <- th$x[th$text == "No"] - 22
  debit_end <- th$x2[th$text == "Debet"]
  credit_end <- th$x2[th$text == "Kredit"]
  dr_cr_split <- (debit_end + credit_end) / 2
  bal_x <- th$x[th$text == "Saldo"] - 15

  body <- table_lines(doc, header = hdr_re, stop = "^J u m l a h")
  stated <- list()
  rows <- list(); cur <- NULL; last_date <- as.Date(NA)

  for (id in body$line_id) {
    w <- line_words(doc, id)
    if (grepl("^Saldo Awal", paste(w$text, collapse = " "))) {
      b <- w[w$x >= bal_x, ]
      stated$opening <- bpd_balance(b$text[1], b$text[2])
      next
    }
    bal_words <- w[w$x >= bal_x, ]
    if (nrow(bal_words) > 0 && is_amount(bal_words$text[1], "id")) {
      if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
      d <- w$text[w$x >= tx_x & w$x < ket_x]
      if (length(d) > 0) last_date <- parse_num_date(d[1])
      amt <- w[w$x >= arsip_x + 60 & w$x < bal_x & is_amount(w$text, "id"), ]
      dr <- sum(parse_amount(amt$text[amt$x2 < dr_cr_split], "id"))
      cr <- sum(parse_amount(amt$text[amt$x2 >= dr_cr_split], "id"))
      cur <- list(date = last_date,
                  desc = paste(w$text[w$x >= ket_x & w$x < code_x], collapse = " "),
                  arsip = paste(w$text[w$x >= arsip_x & w$x < arsip_x + 60], collapse = " "),
                  debit = dr, credit = cr,
                  balance = bpd_balance(bal_words$text[1], if (nrow(bal_words) > 1) bal_words$text[2] else "K"),
                  page = w$page[1])
    } else if (!is.null(cur)) {
      k <- paste(w$text[w$x >= ket_x & w$x < code_x], collapse = " ")
      if (nchar(k)) cur$desc <- c(cur$desc, k)
      a <- paste(w$text[w$x >= arsip_x], collapse = " ")
      if (nchar(a)) cur$arsip <- paste(cur$arsip, a)
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
  rows <- lapply(rows, function(r) { r$desc <- c(r$desc, r$arsip); r })

  # "J u m l a h : <debit> <kredit> <saldo> K"
  j <- which(grepl("^J u m l a h", L$text))
  if (length(j) > 0) {
    w <- line_words(doc, L$line_id[j[length(j)]])
    a <- w[is_amount(w$text, "id"), ]
    if (nrow(a) >= 3) {
      stated$total_debit <- parse_amount(a$text[1], "id")
      stated$total_credit <- parse_amount(a$text[2], "id")
      stated$closing <- bpd_balance(a$text[3], tail(w$text, 1))
    }
  }

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "BPD_DIY", name = "BPD DIY (Bank BPD DI Yogyakarta)", variant = "Rekening Koran",
              description = "Rekening koran BPD DIY: Tgl Val, Tgl Tx, Keterangan, Debet, Kredit, Saldo (format angka 1.234,56)",
              detect = bpd_diy_detect, parse = bpd_diy_parse)
