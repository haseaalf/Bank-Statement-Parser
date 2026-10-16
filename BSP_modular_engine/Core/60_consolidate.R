#' Consolidation of several statements of the same account

#' Grouping key: bank + account number + currency. Statements without an
#' account number are never merged with others.
group_key <- function(bank_id, meta, file_name) {
  if (is.na(meta$account_no) || meta$account_no == "") return(paste0(bank_id, "|?|", file_name))
  paste(bank_id, meta$account_no, meta$currency, sep = "|")
}

statement_start <- function(st) {
  if (!is.na(st$meta$period_start)) return(st$meta$period_start)
  if (nrow(st$txns) > 0) return(min(st$txns$date, na.rm = TRUE))
  as.Date(NA)
}

statement_end <- function(st) {
  if (!is.na(st$meta$period_end)) return(st$meta$period_end)
  if (nrow(st$txns) > 0) return(max(st$txns$date, na.rm = TRUE))
  as.Date(NA)
}

#' Merge statements of one account.
#' @param items list of list(file_name, st) where st is complete_balances()'d
#' @return list(st = merged statement, findings = cross-file findings, files = used file names)
consolidate_statements <- function(items, tol = 0.01) {
  f <- data.frame(file_name = character(), check = character(), severity = character(),
                  message = character(), stringsAsFactors = FALSE)
  add <- function(file, check, sev, msg) {
    f <<- rbind(f, data.frame(file_name = file, check = check, severity = sev,
                              message = msg, stringsAsFactors = FALSE))
  }

  starts <- vapply(items, function(x) as.numeric(statement_start(x$st)), 0)
  items <- items[order(starts)]

  # Duplicate statements (same period and same transactions)
  keep <- rep(TRUE, length(items))
  for (i in seq_along(items)) for (j in seq_len(i - 1)) {
    if (!keep[j]) next
    a <- items[[i]]$st; b <- items[[j]]$st
    if (identical(statement_start(a), statement_start(b)) && identical(statement_end(a), statement_end(b)) &&
        nrow(a$txns) == nrow(b$txns) && isTRUE(all.equal(a$txns$balance, b$txns$balance))) {
      keep[i] <- FALSE
      add(items[[i]]$file_name, "duplicate", "WARN",
          paste0("Duplikat dari ", items[[j]]$file_name, " - tidak ikut digabung"))
      break
    }
  }
  items <- items[keep]

  # Continuity, gaps and overlaps between consecutive periods
  for (i in seq_along(items)[-1]) {
    prev <- items[[i - 1]]; cur <- items[[i]]
    pe <- statement_end(prev$st); cs <- statement_start(cur$st)
    if (!is.na(pe) && !is.na(cs)) {
      if (cs <= pe) add(cur$file_name, "overlap", "WARN",
                        paste0("Periode tumpang tindih dengan ", prev$file_name))
      else if (cs > pe + 1) add(cur$file_name, "gap", "WARN",
                                paste0("Ada periode yang hilang antara ", format(pe, "%d/%m/%Y"),
                                       " dan ", format(cs, "%d/%m/%Y")))
    }
    prev_close <- tail(prev$st$txns$balance, 1)
    if (length(prev_close) == 0) prev_close <- prev$st$opening$value
    cur_open <- cur$st$opening$value
    if (!is.na(prev_close) && !is.na(cur_open) && abs(prev_close - cur_open) > tol) {
      add(cur$file_name, "continuity", "ERROR",
          paste0("Saldo awal ", fmt_num(cur_open), " tidak sama dengan saldo akhir periode sebelumnya ",
                 fmt_num(prev_close), " (", prev$file_name, ")"))
    }
  }

  first <- items[[1]]$st; last <- items[[length(items)]]$st
  txns <- do.call(rbind, lapply(items, function(x) x$st$txns))
  sum_or_na <- function(field) {
    v <- vapply(items, function(x) as.numeric(x$st$stated[[field]]), 0)
    if (any(is.na(v))) NA else sum(v)
  }
  st <- new_statement(txns,
                      meta = modifyList(first$meta, list(period_start = statement_start(first),
                                                         period_end = statement_end(last))),
                      stated = list(opening = first$opening$value, closing = last$stated$closing,
                                    total_debit = sum_or_na("total_debit"),
                                    total_credit = sum_or_na("total_credit"),
                                    n_debit = sum_or_na("n_debit"), n_credit = sum_or_na("n_credit")))
  st$txns$balance_source <- txns$balance_source
  st$opening <- first$opening

  list(st = st, findings = f, files = vapply(items, `[[`, "", "file_name"))
}
