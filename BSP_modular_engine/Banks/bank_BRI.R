#' BRI - "Laporan Transaksi Finansial / Statement of Financial Transaction" (BRISIM)
#'
#' Row layout : dd/mm/yy hh:mm:ss | description (1-3 lines) | teller | debit | credit | balance
#' Summary    : last transaction page shows Opening Balance, Total Debit,
#'              Total Credit, Closing Balance

bri_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("LAPORAN TRANSAKSI FINANSIAL", "STATEMENT OF FINANCIAL TRANSACTION",
                  "BRISIM", "Teller", "Unit Kerja", "BANK BRI"),
                required = "TRANSAKSI FINANSIAL")
}

bri_parse <- function(doc) {
  L <- doc$lines
  txt <- doc_all_text(doc)

  # ---- Header / meta -------------------------------------------------------
  period <- regmatches(txt, regexec("Periode Transaksi\\s*:\\s*(\\d\\d/\\d\\d/\\d\\d)\\s*-\\s*(\\d\\d/\\d\\d/\\d\\d)", txt))[[1]]
  meta <- list(account_no = grab(txt, "No\\. Rekening\\s*:\\s*(\\d+)"),
               currency = grab(txt, "Valuta\\s*:\\s*([A-Z]{3})"),
               period_start = if (length(period) == 3) parse_num_date(period[2]) else as.Date(NA),
               period_end = if (length(period) == 3) parse_num_date(period[3]) else as.Date(NA))
  to_line <- which(grepl("^Kepada", L$text))[1]
  if (!is.na(to_line)) {
    # the name is the first line below "Kepada Yth." in the left column
    for (k in seq(to_line + 1, length.out = 4)) {
      w <- line_words(doc, L$line_id[k])
      if (any(w$x < 300)) { meta$account_name <- paste(w$text[w$x < 300], collapse = " "); break }
    }
  }

  # ---- Summary block ------------------------------------------------------
  stated <- list()
  sum_hdr <- which(grepl("Opening Balance", L$text) & grepl("Closing Balance", L$text))
  if (length(sum_hdr) > 0) {
    v <- strsplit(L$text[sum_hdr[1] + 1], " ")[[1]]
    v <- parse_amount(v[is_amount(v)])
    if (length(v) == 4) stated <- list(opening = v[1], total_debit = v[2],
                                       total_credit = v[3], closing = v[4])
  }

  # ---- Transactions ---------------------------------------------------------
  rows <- list()
  cur <- NULL
  date_re <- "^\\d\\d/\\d\\d/\\d\\d \\d\\d:\\d\\d:\\d\\d"

  for (p in seq_len(doc$n_pages)) {
    pl <- L[L$page == p, ]
    hdr <- which(grepl("Transaction Date", pl$text) & grepl("Balance", pl$text))
    if (length(hdr) == 0) next
    hw <- line_words(doc, pl$line_id[hdr[1]])
    # description starts after the date+time; continuation lines are indented past the date
    desc_x <- min(hw$x) + 40
    teller_x <- min(hw$x[hw$text %in% c("User", "Teller")], na.rm = TRUE) - 2

    for (k in seq(hdr[1] + 1, length.out = max(0, nrow(pl) - hdr[1]))) {
      line <- pl[k, ]
      if (grepl("Opening Balance|^Saldo Awal|Created By", line$text)) break
      w <- line_words(doc, line$line_id)

      if (grepl(date_re, line$text)) {
        if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
        amt <- which(is_amount(w$text) & w$x >= teller_x)
        amt <- tail(amt, 3)
        vals <- parse_amount(w$text[amt])
        desc_words <- seq_len(nrow(w)) > 2 & w$x < teller_x
        cur <- list(date = parse_num_date(w$text[1]),
                    desc = paste(w$text[desc_words], collapse = " "),
                    debit = if (length(vals) == 3) vals[1] else NA,
                    credit = if (length(vals) == 3) vals[2] else NA,
                    balance = if (length(vals) == 3) vals[3] else NA,
                    page = p)
      } else if (!is.null(cur) && all(w$x >= desc_x & w$x < teller_x)) {
        cur$desc <- c(cur$desc, paste(w$text, collapse = " "))
      }
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur

  txns <- data.frame(date = as.Date(vapply(rows, function(r) as.numeric(r$date), 0), origin = "1970-01-01"),
                     description = vapply(rows, function(r) join_desc(r$desc), ""),
                     debit = vapply(rows, function(r) r$debit, 0),
                     credit = vapply(rows, function(r) r$credit, 0),
                     balance = vapply(rows, function(r) r$balance, 0),
                     page = vapply(rows, function(r) r$page, 0L),
                     stringsAsFactors = FALSE)
  new_statement(txns, meta, stated)
}

register_bank(id = "BRI", name = "Bank Rakyat Indonesia (BRI)",
              variant = "Laporan Transaksi Finansial",
              description = "Rekening koran BRISIM: kolom Teller, Debet, Kredit, Saldo",
              detect = bri_detect, parse = bri_parse)
