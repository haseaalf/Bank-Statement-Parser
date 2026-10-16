#' Maybank Indonesia - Regional Cash Management "Account Report" (CA)
#'
#' Row layout : Transaction Date | Time | Posting Date | Processing Time | Description |
#'              Transaction Ref | Debit | Credit | Source Code | Teller ID |
#'              Branch/Channel | Transaction Code | End Balance
#'              Narrow columns wrap onto following lines (also amounts: "6,560,000.0" + "0").
#' Summary    : Begin / End Balance, Total Debit / Total Credit (= number of records)

maybank_detect <- function(doc) {
  keyword_score(doc_all_text(doc),
                c("REGIONAL CASH MANAGEMENT", "Account Report", "Begin Balance", "End Balance",
                  "Source Code", "Teller ID", "Total no\\. of Records"),
                required = "Begin Balance")
}

maybank_parse <- function(doc) {
  L <- doc$lines
  txt <- doc_all_text(doc)
  head_w <- doc$words[doc$words$page == 1 & doc$words$y < 130, ]
  acc <- head_w$text[grepl("^\\d{8,}$", head_w$text)][1]
  name_line <- which(grepl("^Account Name", L$text))[1]
  nw <- line_words(doc, L$line_id[name_line])
  meta <- list(account_no = acc,
               account_name = paste(nw$text[nw$x > 120 & nw$x < 600], collapse = " "),
               currency = grab(txt, "Currency\\s*:\\s*([A-Z]{3})"),
               period_start = parse_num_date(grab(txt, "Start Date\\s*:\\s*(\\S+)")),
               period_end = parse_num_date(grab(txt, "End Date\\s*:\\s*(\\S+)")))
  stated <- list(opening = parse_amount(grab(txt, "Begin Balance\\s*:\\s*([0-9.,-]+)")),
                 closing = parse_amount(grab(txt, "End Balance\\s*:\\s*([0-9.,-]+)")),
                 n_debit = as.integer(grab(txt, "Total Debit\\s*:\\s*(\\d+)")),
                 n_credit = as.integer(grab(txt, "Total Credit\\s*:\\s*(\\d+)")))

  hdr_re <- "^Transaction Transaction Posting Date"
  th <- line_words(doc, L$line_id[which(grepl(hdr_re, L$text))[1]])
  hx <- function(label, k = 1) th$x[th$text == label][k]
  post_x  <- hx("Posting") - 5
  proc_x  <- hx("Processing") - 2
  desc_x  <- hx("Transaction", 3) - 2
  ref_x   <- hx("Transaction", 4) - 2
  debit_x <- hx("Debit") - 30
  credit_x <- hx("Credit") - 21
  source_x <- hx("Source") - 2
  bal_x <- hx("End") - 40

  body <- table_lines(doc, header = hdr_re, stop = "^Total no\\. of Records")
  body <- body[!grepl("^Date Time Time Description", body$text), ]   # 2nd header line

  is_anchor <- vapply(body$line_id, function(id) {
    w <- line_words(doc, id)
    any(w$x >= bal_x & is_amount(w$text))
  }, logical(1))
  owner <- cumsum(is_anchor)                 # following lines belong to the anchor above

  rows <- lapply(seq_len(sum(is_anchor)), function(a) {
    blk <- body[owner == a, ]
    col <- function(x0, x1, sep = "") {
      vapply(blk$line_id, function(id) {
        w <- line_words(doc, id)
        paste(w$text[w$x >= x0 & w$x < x1], collapse = " ")
      }, "")
    }
    col_join <- function(x0, x1) paste(col(x0, x1), collapse = "")   # wrapped values
    trx_date <- parse_text_date(col(0, post_x)[1])
    post_date <- parse_text_date(col(post_x, proc_x)[1])
    desc <- col(desc_x, ref_x)
    desc <- desc[nchar(desc) > 0]
    # a reference wrapped over several lines ("20260326BNI" / "AIDJA010O010" / "0238349") is one token
    code <- grepl("^[A-Z0-9]+$", desc) & grepl("\\d", desc)
    if (length(desc) > 1) for (k in rev(seq_along(desc))[-length(desc)]) {
      if (code[k] && code[k - 1]) { desc[k - 1] <- paste0(desc[k - 1], desc[k]); desc[k] <- "" }
    }
    ref <- col_join(ref_x, debit_x)
    aw <- line_words(doc, blk$line_id[1])
    list(date = if (!is.na(trx_date)) trx_date else post_date,
         desc = c(desc[nchar(desc) > 0], ref),
         debit = parse_amount(gsub(" ", "", col_join(debit_x, credit_x))),
         credit = parse_amount(gsub(" ", "", col_join(credit_x, source_x))),
         balance = parse_amount(aw$text[aw$x >= bal_x & is_amount(aw$text)][1]),
         page = blk$page[1])
  })

  new_statement(rows_to_txns(rows), meta, stated)
}

register_bank(id = "MAYBANK", name = "Maybank Indonesia", variant = "Cash Management Account Report",
              description = "Account Report Maybank (Regional Cash Management): Debit, Credit, End Balance",
              detect = maybank_detect, parse = maybank_parse)
