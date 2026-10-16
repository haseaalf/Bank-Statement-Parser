#' Bank Syariah Indonesia (BSI) - Account Statement (internet / cash management)
#'
#' Row layout : yyyy-mm-dd hh:mm:ss | FT Number | Description | Currency | Amount | DB/CR | Balance
#' Summary    : Opening / Closing Balance, Total Debit/Credit Amount and Record

bsi_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("Account Statement", "FT Number", "Opening Balance", "Total Debit Record",
                  "Total Credit Record"),
                required = "FT Number")
}

bsi_parse <- function(doc) {
  txt <- doc_all_text(doc)
  L <- doc$lines
  amt_after <- function(label) parse_amount(grab(txt, paste0(label, "\\s*:\\s*(?:[A-Z]{3}\\s*)?([0-9.,]+)")))

  period <- grab(txt, "Date\\s*:\\s*(\\d{1,2} \\w+ \\d{4}\\s*-\\s*\\d{1,2} \\w+ \\d{4})")
  pd <- strsplit(period, "\\s*-\\s*")[[1]]
  meta <- list(account_no = grab(txt, "Account\\s*:\\s*(\\d+)"),
               account_name = squish(sub("/\\s*[A-Z]{3}\\s*$", "", L$text[2])),
               currency = grab(L$text[2], "/\\s*([A-Z]{3})\\s*$"),
               period_start = parse_text_date(pd[1]), period_end = parse_text_date(pd[2]))
  stated <- list(opening = amt_after("Opening Balance"), closing = amt_after("Closing Balance"),
                 total_debit = amt_after("Total Debit Amount"), total_credit = amt_after("Total Credit Amount"),
                 n_debit = as.integer(grab(txt, "Total Debit Record\\s*:\\s*(\\d+)")),
                 n_credit = as.integer(grab(txt, "Total Credit Record\\s*:\\s*(\\d+)")))

  hdr_re <- "^Date FT Number Description"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  desc_x <- th$x[th$text == "Description"] - 5
  cur_x <- th$x[th$text == "Currency"] - 5
  ft_x <- th$x[th$text == "FT"] - 5

  body <- table_lines(doc, header = hdr_re, stop = "^Page \\d+/\\d+", header_optional = TRUE)
  rows <- list(); cur <- NULL
  for (id in body$line_id) {
    w <- line_words(doc, id)
    if (grepl("^\\d{4}-\\d{2}-\\d{2}$", w$text[1])) {
      if (!is.null(cur)) rows[[length(rows) + 1]] <- cur
      right <- w[w$x >= cur_x, ]
      tokens <- paste(right$text, collapse = " ")
      # "IDR 10,000.00 DB 1,787,313.77" or "IDR 10,000.00DB 1,787,313.77"
      m <- regmatches(tokens, regexec("([0-9.,]+)\\s*(DB|CR)\\s+(-?[0-9.,]+)", tokens))[[1]]
      amount <- parse_amount(m[2])
      cur <- list(date = as.Date(w$text[1]),
                  desc = c(paste(w$text[w$x >= desc_x & w$x < cur_x], collapse = " "),
                           paste(w$text[w$x >= ft_x & w$x < desc_x], collapse = " ")),
                  debit = if (identical(m[3], "DB")) amount else 0,
                  credit = if (identical(m[3], "CR")) amount else 0,
                  balance = parse_amount(m[4]), page = w$page[1])
    } else if (!is.null(cur) && all(w$x >= desc_x & w$x < cur_x)) {
      # wrapped description: keep before the FT Number line
      n <- length(cur$desc)
      cur$desc <- c(cur$desc[-n], paste(w$text, collapse = " "), cur$desc[n])
    }
  }
  if (!is.null(cur)) rows[[length(rows) + 1]] <- cur

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "BSI", name = "Bank Syariah Indonesia (BSI)", variant = "Account Statement",
              description = "Account Statement BSI: Date, FT Number, Description, Amount DB/CR, Balance",
              detect = bsi_detect, parse = bsi_parse)
