#' CIMB Niaga - Account statement (BizChannel)
#'
#' Row layout : No | Post Date (mm/dd/yy hh:mm) | Eff Date | Cheque No | Description |
#'              Debit | Credit | Balance | Transaction Ref No
#' Order      : newest first - rows are put back in chronological order
#' Summary    : Total Debit / Total Credit only (no opening / closing balance)

cimb_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("Post Date", "Eff Date", "Cheque No", "Transaction Ref No", "Account Type",
                  "Total Debit", "Total Credit"),
                required = "Transaction Ref No")
}

cimb_parse <- function(doc) {
  txt <- doc_all_text(doc)
  L <- doc$lines
  value_of <- function(label) {
    i <- which(grepl(paste0("^", label, "\\b"), L$text))[1]
    if (is.na(i)) return(NA_character_)
    w <- line_words(doc, L$line_id[i])
    paste(w$text[w$x >= 150], collapse = " ")
  }
  period <- grab(txt, "Period\\s*:\\s*(\\S+)\\s*-\\s*(\\S+)", 1)
  period_end <- grab(txt, "Period\\s*:\\s*(\\S+)\\s*-\\s*(\\S+)", 2)
  meta <- list(account_no = value_of("Account Number"), account_name = value_of("Account Name"),
               currency = value_of("Currency"),
               period_start = parse_text_date(period), period_end = parse_text_date(period_end))
  stated <- list(total_debit = parse_amount(value_of("Total Debit")),
                 total_credit = parse_amount(value_of("Total Credit")))

  hdr_re <- "^No\\. Post Date"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  desc_x <- th$x[th$text == "Cheque"] + 40           # description starts after the cheque column
  cheque_x <- th$x[th$text == "Cheque"] - 5
  ref_x <- th$x[th$text == "Transaction"] - 5

  body <- table_lines(doc, header = hdr_re, header_optional = TRUE)
  rows <- list(); cur <- NULL
  for (id in body$line_id) {
    w <- line_words(doc, id)
    if (grepl("^\\d+$", w$text[1]) && nrow(w) > 2 && grepl("^\\d\\d/\\d\\d/\\d\\d$", w$text[2])) {
      if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
      amt <- tail(which(is_amount(w$text) & w$x < ref_x & w$x > desc_x), 3)
      v <- parse_amount(w$text[amt])
      first_amt_x <- min(w$x[amt])
      cheque <- paste(w$text[w$x >= cheque_x & w$x < desc_x], collapse = " ")
      ref <- paste(w$text[w$x >= ref_x], collapse = " ")
      cur <- list(date = parse_num_date(w$text[2], order = "mdy"),
                  desc = c(paste(w$text[w$x >= desc_x & w$x < first_amt_x], collapse = " "),
                           cheque, ref),
                  debit = v[1], credit = v[2], balance = v[3], page = w$page[1])
    } else if (!is.null(cur) && nrow(w) > 0 && all(w$x >= desc_x & w$x < ref_x) && !any(is_amount(w$text))) {
      n <- length(cur$desc)
      cur$desc <- c(cur$desc[1:(n - 2)], paste(w$text, collapse = " "), cur$desc[(n - 1):n])
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur

  txns <- fix_order_by_chain(rows_to_txns(rows))
  new_statement(txns, meta, stated)
}

register_bank(id = "CIMB", name = "CIMB Niaga", variant = "BizChannel Account Statement",
              description = "Statement CIMB Niaga: No, Post Date (MM/DD/YY), Eff Date, Description, Debit, Credit, Balance, Transaction Ref No",
              detect = cimb_detect, parse = cimb_parse)
