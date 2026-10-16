#' BNI - Account Statement (BNIDirect / e-channel download)
#'
#' Row layout : Posting Date | Effective Date | Branch | Journal | Transaction Description |
#'              Amount | DB/CR (K = credit, D = debit) | Balance
#'              Every cell is vertically centred and spans several lines, so rows
#'              are separated with the table rules drawn in the PDF.
#'              Glyphs are drawn twice (handled by drop_ghost_words()).
#' Summary    : "Ledger Balance:" (opening), "Ending Balance :", Total Debet / Credit
#'              (count + amount)

bni_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("Account Information", "ACCOUNT STATEMENT", "Posting Date", "Effective Date",
                  "Journal", "DB/CR", "Ledger Balance"),
                required = "Effective Date.*Branch.*Journal")
}

bni_parse <- function(doc) {
  L <- doc$lines
  txt <- gsub("\\s+", " ", doc_all_text(doc))

  acc <- grab(txt, "Account No\\.\\s*:\\s*(\\d+)\\s*/\\s*(.*?)\\s*\\(?([A-Z]{3})\\)")
  name_cur <- regmatches(txt, regexec("Account No\\.\\s*:\\s*\\d+\\s*/\\s*(.*?)\\s*\\(([A-Z]{3})\\)", txt, perl = TRUE))[[1]]
  per <- regmatches(txt, regexec("Period\\s*:\\s*(\\S+)\\s*-\\s*(\\S+)", txt))[[1]]
  meta <- list(account_no = acc,
               account_name = if (length(name_cur) == 3) name_cur[2] else NA,
               currency = if (length(name_cur) == 3) name_cur[3] else "IDR",
               period_start = parse_text_date(per[2]), period_end = parse_text_date(per[3]))
  tot_db <- regmatches(txt, regexec("Total Debet\\s*:\\s*(\\d+)\\s+([0-9.,]+)", txt))[[1]]
  tot_cr <- regmatches(txt, regexec("Total Credit\\s*:\\s*(\\d+)\\s+([0-9.,]+)", txt))[[1]]
  stated <- list(opening = parse_amount(grab(txt, "Ledger Balance\\s*:\\s*([0-9.,-]+)")),
                 closing = parse_amount(grab(txt, "Ending Balance\\s*:\\s*([0-9.,-]+)")),
                 n_debit = as.integer(tot_db[2]), total_debit = parse_amount(tot_db[3]),
                 n_credit = as.integer(tot_cr[2]), total_credit = parse_amount(tot_cr[3]))

  hdr_re <- "^Posting Date Effective Date"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  desc_x <- th$x[th$text == "Transaction"] - 22
  amt_x <- th$x[th$text == "Amount"] - 20
  drcr_x <- th$x[th$text == "DB/CR"] - 1
  bal_x <- th$x[th$text == "Balance"] - 15

  body <- table_lines(doc, header = hdr_re, stop = "^Ending Balance")
  body$anchor <- grepl("^\\d\\d/\\d\\d/\\d{4} ", body$text)

  rows <- list()
  for (p in unique(body$page)) {
    pb <- body[body$page == p, ]
    rules <- table_row_rules(doc, p, desc_x + 5, amt_x - 5)
    hdr_y <- L$y[L$page == p & grepl(hdr_re, L$text)][1]
    rules <- rules[rules > hdr_y + 5]
    if (length(rules) >= 2) {
      cell <- findInterval(pb$y + 2, rules)
    } else {
      a <- pb[pb$anchor, ]
      cell <- if (nrow(a)) nearest_anchor(pb, a) else rep(0L, nrow(pb))
    }
    for (cid in unique(cell)) {
      blk <- pb[cell == cid, ]
      blk <- blk[order(blk$y), ]
      desc <- vapply(blk$line_id, function(id) {
        w <- line_words(doc, id)
        paste(w$text[w$x >= desc_x & w$x < amt_x], collapse = " ")
      }, "")
      desc <- desc[nchar(desc) > 0]
      if (!any(blk$anchor)) {
        # a row continued from the previous page (cells before the first anchor are skipped)
        first_cell <- cid == cell[1]
        if (first_cell && length(rows) > 0 && rows[[length(rows)]]$page < p) {
          rows[[length(rows)]]$desc <- c(rows[[length(rows)]]$desc, desc)
        }
        next
      }
      aw <- line_words(doc, blk$line_id[blk$anchor][1])
      a <- aw$text[aw$x >= amt_x & aw$x < drcr_x & is_amount(aw$text)]
      mark <- aw$text[aw$x >= drcr_x & aw$x < bal_x]
      b <- aw$text[aw$x >= bal_x & is_amount(aw$text)]
      amount <- parse_amount(a[1])
      rows[[length(rows) + 1]] <- list(
        date = parse_num_date(aw$text[1]), desc = desc,
        debit = if (any(mark == "D")) amount else 0,
        credit = if (any(mark == "K")) amount else 0,
        balance = if (length(b)) parse_amount(b[1]) else NA, page = p)
    }
  }

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "BNI", name = "Bank Negara Indonesia (BNI)", variant = "Account Statement",
              description = "Account Statement BNI: Posting Date, Effective Date, Branch, Journal, Description, Amount, DB/CR, Balance",
              detect = bni_detect, parse = bni_parse)
