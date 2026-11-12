#' Bank Mandiri - e-Statement Tabungan (Livin' by Mandiri, PDF with password)
#'
#' Row layout : No | Tanggal/Date ("01 Oct 2026" + "11:37:18 WIB" on the lines around) |
#'              Keterangan/Remarks (1-4 lines above and below) | Nominal (+/- amount) | Saldo
#'              Amounts in Indonesian format (1.234.567,89). The line with No + amount is the
#'              anchor; the other lines of a transaction are vertically centred around it.
#' Summary    : page 1 - Saldo Awal/Initial Balance, Dana Masuk/Incoming, Dana Keluar/Outgoing,
#'              Saldo Akhir/Closing Balance

mandiri_estatement_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("e-Statement", "Menara Mandiri", "Saldo Awal/Initial Balance", "Dana Masuk/Incoming",
                  "Dana Keluar/Outgoing", "Nominal \\(IDR\\)", "Bank Mandiri"),
                required = "Saldo Awal/Initial Balance")
}

mandiri_estatement_parse <- function(doc) {
  L <- doc$lines
  txt <- gsub("\\s+", " ", doc_all_text(doc))

  # ---- Meta (page 1 header)
  value_after <- function(label_re, x_from, x_to = Inf) {
    i <- which(grepl(label_re, L$text))[1]
    if (is.na(i)) return(NA_character_)
    w <- line_words(doc, L$line_id[i])
    paste(w$text[w$x >= x_from & w$x < x_to], collapse = " ")
  }
  name_line <- line_words(doc, L$line_id[which(grepl("^Nama/Name", L$text))[1]])
  colon <- name_line$x[name_line$text == ":"][1]
  name_end <- name_line$x[grepl("^Periode", name_line$text)][1]
  period <- grab(txt, "Periode/Period\\s*:\\s*(\\d{1,2} \\w+ \\d{4}\\s*-\\s*\\d{1,2} \\w+ \\d{4})")
  pd <- strsplit(period, "\\s*-\\s*")[[1]]
  meta <- list(account_no = grab(txt, "Account Number\\s*:\\s*(\\d+)"),
               account_name = paste(name_line$text[name_line$x > colon & name_line$x < name_end], collapse = " "),
               currency = grab(txt, "Uang/Currency\\s*:\\s*([A-Z]{3})"),
               period_start = parse_text_date(pd[1]), period_end = parse_text_date(pd[2]))
  amount_after <- function(label) parse_amount(grab(txt, paste0(label, "\\s*:\\s*[+-]?\\s*([0-9.,]+)")), "id")
  stated <- list(opening = amount_after("Saldo Awal/Initial Balance"),
                 closing = amount_after("Saldo Akhir/Closing Balance"),
                 total_credit = amount_after("Dana Masuk/Incoming Transactions"),
                 total_debit = amount_after("Dana Keluar/Outgoing Transactions"))

  # ---- Columns from the (English) table header
  hdr_re <- "^No Date Remarks"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  date_x <- th$x[th$text == "Date"] - 5
  rem_x <- th$x[th$text == "Remarks"] - 5
  amt_x <- th$x[th$text == "Amount"] - 40
  bal_x <- th$x[th$text == "Balance"] - 30

  body <- table_lines(doc, header = hdr_re, stop = "^PT Bank Mandiri|^Disclaimer")
  body <- body[!grepl("^Mandiri Call", body$text), ]
  signed <- "^[+-][0-9.]+,[0-9]{2}$"
  body$anchor <- vapply(body$line_id, function(id) {
    w <- line_words(doc, id)
    grepl("^\\d+$", w$text[1]) && w$x[1] < date_x && any(grepl(signed, w$text[w$x >= amt_x & w$x < bal_x]))
  }, logical(1))

  rows <- list()
  for (p in unique(body$page)) {
    pb <- body[body$page == p, ]
    a <- pb[pb$anchor, ]
    if (nrow(a) == 0) {
      # page holding only the tail of the previous transaction
      if (length(rows) > 0) rows[[length(rows)]]$lines <- rbind(rows[[length(rows)]]$lines, pb)
      next
    }
    owner <- nearest_anchor(pb, a)
    # lines above the first anchor that are far from it continue the last transaction of the previous page
    carry <- pb$y < a$y[1] - 20
    if (any(carry) && length(rows) > 0) rows[[length(rows)]]$lines <- rbind(rows[[length(rows)]]$lines, pb[carry, ])
    for (k in seq_len(nrow(a))) {
      rows[[length(rows) + 1]] <- list(anchor = a$line_id[k], page = p, lines = pb[owner == k & !carry, ])
    }
  }

  txns <- lapply(rows, function(r) {
    blk <- r$lines[order(r$lines$page, r$lines$y), ]
    date_txt <- character(); remarks <- character()
    for (id in blk$line_id) {
      w <- line_words(doc, id)
      date_txt <- c(date_txt, w$text[w$x >= date_x & w$x < rem_x])
      rm <- paste(w$text[w$x >= rem_x & w$x < amt_x], collapse = " ")
      if (nchar(rm) && rm != "-") remarks <- c(remarks, rm)
    }
    aw <- line_words(doc, r$anchor)
    amt <- aw$text[aw$x >= amt_x & aw$x < bal_x & grepl(signed, aw$text)][1]
    v <- parse_amount(sub("^[+-]", "", amt), "id")
    bal <- aw$text[aw$x >= bal_x & is_amount(aw$text, "id")][1]
    list(date = parse_text_date(paste(date_txt, collapse = " ")), desc = remarks,
         debit = if (startsWith(amt, "-")) v else 0, credit = if (startsWith(amt, "+")) v else 0,
         balance = parse_amount(bal, "id"), page = r$page)
  })

  new_statement(rows_to_txns(txns), meta, stated)
}

register_bank(id = "MANDIRI", name = "Bank Mandiri", variant = "e-Statement Tabungan (Livin')",
              description = "e-Statement tabungan Mandiri (Livin'): No, Tanggal, Keterangan, Nominal (+/-), Saldo; PDF berpassword",
              detect = mandiri_estatement_detect, parse = mandiri_estatement_parse)
