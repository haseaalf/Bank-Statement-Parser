#' Bank Mandiri - Account Statement (Kopra by Mandiri / Mandiri Cash Management)
#'
#' Row layout : the line holding Debit / Credit / Balance is the anchor; date
#'              ("13 Mar 2026," + time) and Remark lines sit above and below it.
#' Summary    : Opening / Closing Balance, No. of Debit / Credit, Total Amount
#'              Debited / Credited

mandiri_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("Account Statement Summary", "Posting Date", "Remark", "Reference No",
                  "Total Amount Debited", "Mandiri"),
                required = "Mandiri")
}

mandiri_parse <- function(doc) {
  L <- doc$lines

  # ---- Summary (label line followed by value line, values aligned under labels)
  value_below <- function(label_re, x_from, x_to) {
    i <- which(grepl(label_re, L$text))[1]
    if (is.na(i)) return(NA_character_)
    w <- line_words(doc, L$line_id[i + 1])
    paste(w$text[w$x >= x_from & w$x < x_to], collapse = " ")
  }
  hdr <- line_words(doc, L$line_id[which(grepl("^Account No\\.", L$text))[1]])
  col2 <- hdr$x[hdr$text == "Account"][2] - 2      # "Account Name" column
  col3 <- hdr$x[hdr$text == "Alias"] - 2

  period <- value_below("^Period", 0, col2)
  pd <- strsplit(period, " - ")[[1]]
  meta <- list(account_no = value_below("^Account No\\.", 0, col2),
               account_name = value_below("^Account No\\.", col2, col3),
               currency = value_below("^Period", col2, col3),
               period_start = parse_text_date(pd[1]),
               period_end = parse_text_date(pd[2]))
  stated <- list(opening = parse_amount(value_below("^Opening Balance", 0, col2)),
                 n_debit = as.integer(value_below("^Opening Balance", col2, col3)),
                 total_debit = parse_amount(value_below("^Opening Balance", col3, Inf)),
                 closing = parse_amount(value_below("^Closing Balance", 0, col2)),
                 n_credit = as.integer(value_below("^Closing Balance", col2, col3)),
                 total_credit = parse_amount(value_below("^Closing Balance", col3, Inf)))

  # ---- Transactions
  body <- table_lines(doc, header = "^Posting Date.*Balance", stop = "For further questions|Page \\d+ of")
  th <- line_words(doc, L$line_id[which(grepl("^Posting Date.*Balance", L$text))[1]])
  remark_x <- th$x[th$text == "Remark"] - 5
  ref_x <- th$x[th$text == "Reference"] - 5

  is_anchor <- vapply(body$line_id, function(id) {
    w <- line_words(doc, id)
    sum(is_amount(w$text) & w$x > ref_x) >= 3
  }, logical(1))
  anchors <- body[is_anchor, ]
  owner <- nearest_anchor(body, anchors)

  rows <- lapply(seq_len(nrow(anchors)), function(a) {
    blk <- body[which(owner == a), ]
    blk <- blk[order(blk$y), ]
    date_txt <- character(); remark <- character(); ref <- character()
    for (id in blk$line_id) {
      w <- line_words(doc, id)
      date_txt <- c(date_txt, w$text[w$x < remark_x])
      r <- w$text[w$x >= remark_x & w$x < ref_x]
      if (length(r)) remark <- c(remark, paste(r, collapse = " "))
    }
    aw <- line_words(doc, anchors$line_id[a])
    amt <- tail(which(is_amount(aw$text) & aw$x > ref_x), 3)
    v <- parse_amount(aw$text[amt])
    refs <- aw$text[aw$x >= ref_x & !(seq_len(nrow(aw)) %in% amt) & aw$text != "-"]
    list(date = parse_text_date(paste(date_txt, collapse = " ")),
         desc = c(remark, refs), debit = v[1], credit = v[2], balance = v[3],
         page = anchors$page[a])
  })

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "MANDIRI", name = "Bank Mandiri", variant = "Account Statement (Kopra)",
              description = "Account Statement Kopra by Mandiri: Posting Date, Remark, Reference No., Debit, Credit, Balance",
              detect = mandiri_detect, parse = mandiri_parse)
