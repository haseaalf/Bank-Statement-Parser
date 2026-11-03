#' Monthly recap ("rekap mutasi") over many accounts / banks / customers
#'
#' For every account (bank + account number + currency) the statements are
#' consolidated (sorted, duplicates dropped) and summarised per calendar month:
#' total debit, total credit and end-of-month balance. Every month row is
#' verified: previous month balance + credit - debit = this month balance.
#' Problems are attached to the month row as warnings (shown next to it).

month_key <- function(d) format(d, "%Y-%m")
month_first <- function(key) as.Date(paste0(key, "-01"))

months_between <- function(from, to) {
  if (is.na(from) || is.na(to)) return(character())
  month_key(seq(month_first(month_key(from)), month_first(month_key(to)), by = "month"))
}

#' @param items list of list(file_name, st, bank) for ONE account (st complete_balances()'d)
#' @return list(bank_id, account_no, account_name, currency, table, warnings)
recap_account <- function(items, months, tol = 0.01) {
  cons <- consolidate_statements(lapply(items, function(x) list(file_name = x$file_name, st = x$st)), tol)
  kept <- items[match(cons$files, vapply(items, `[[`, "", "file_name"))]
  kept <- kept[order(vapply(kept, function(x) as.numeric(statement_start(x$st)), 0))]
  t <- cons$st$txns
  t_month <- month_key(t$date)

  # which statement covers each month (first one wins)
  cover <- setNames(rep(NA_integer_, length(months)), months)
  for (k in seq_along(kept)) {
    m <- months_between(statement_start(kept[[k]]$st), statement_end(kept[[k]]$st))
    m <- intersect(m, months)
    cover[m][is.na(cover[m])] <- k
  }

  own <- c(months[!is.na(cover)], unique(t_month))
  acct_first <- min(own); acct_last <- max(own)
  warn <- setNames(vector("list", length(months)), months)
  add_warn <- function(m, txt) if (!is.na(m) && m %in% months) warn[[m]] <<- c(warn[[m]], txt)

  # cross-file findings (duplicate / overlap / gap / continuity) on the month the file starts
  for (i in seq_len(nrow(cons$findings))) {
    f <- cons$findings[i, ]
    src <- items[[match(f$file_name, vapply(items, `[[`, "", "file_name"))]]
    add_warn(month_key(statement_start(src$st)), f$message)
  }

  rows <- data.frame(month = months, debit = NA_real_, credit = NA_real_, saldo = NA_real_,
                     has_data = FALSE, n_tx = 0L, stringsAsFactors = FALSE)
  prev <- NA_real_
  prev_missing <- FALSE
  seen_data <- FALSE
  for (i in seq_along(months)) {
    m <- months[i]
    sel <- which(t_month == m)
    covered <- !is.na(cover[m])
    if (!covered && length(sel) == 0) {
      # only a gap inside the account's own period is a problem; months before its first or
      # after its last statement are simply outside what was uploaded for this account
      if (m > acct_first && m < acct_last) add_warn(m, "Rekening koran bulan ini tidak ada / tidak di-upload")
      prev_missing <- seen_data
      next
    }
    if (is.na(prev)) {
      # first available month: start from the opening balance of the statement covering it
      k <- if (covered) cover[m] else 1L
      prev <- kept[[k]]$st$opening$value
    } else if (prev_missing) {
      add_warn(m, "Bulan sebelumnya tidak ada - urutan bulan tidak bersambung")
      k <- if (covered) cover[m] else NA
      if (!is.na(k)) prev <- kept[[k]]$st$opening$value
    }
    debit <- round(sum(t$debit[sel]), 2)
    credit <- round(sum(t$credit[sel]), 2)
    saldo <- if (length(sel) > 0) t$balance[max(sel)] else prev
    if (!is.na(prev) && !is.na(saldo)) {
      expected <- round(prev + credit - debit, 2)
      if (abs(expected - saldo) > tol) {
        add_warn(m, paste0("Saldo bulan lalu + kredit - debit = ", fmt_num(expected), ", saldo akhir = ",
                           fmt_num(saldo), " (selisih ", fmt_num(saldo - expected), ")"))
      }
    }
    rows$debit[i] <- debit; rows$credit[i] <- credit; rows$saldo[i] <- saldo
    rows$has_data[i] <- TRUE; rows$n_tx[i] <- length(sel)
    prev <- saldo
    prev_missing <- FALSE
    seen_data <- TRUE
  }
  rows$warning <- vapply(warn, function(w) paste(unique(w), collapse = "; "), "")

  meta <- kept[[1]]$st$meta
  list(bank_id = kept[[1]]$bank$id, account_no = meta$account_no, account_name = meta$account_name,
       currency = meta$currency, table = rows, files = cons$files)
}

#' Build the recap for all parsed statements.
#' @param parsed list of list(ok, bank, st, file_path)
build_recap <- function(parsed, tol = 0.01) {
  ok <- Filter(function(r) isTRUE(r$ok), parsed)
  if (length(ok) == 0) return(NULL)
  starts <- do.call(c, lapply(ok, function(r) statement_start(r$st)))
  ends <- do.call(c, lapply(ok, function(r) statement_end(r$st)))
  months <- months_between(min(starts, na.rm = TRUE), max(ends, na.rm = TRUE))

  keys <- vapply(ok, function(r) group_key(r$bank$id, r$st$meta, file_label(r$file_path)), "")
  accounts <- lapply(unique(keys), function(k) {
    grp <- ok[keys == k]
    recap_account(lapply(grp, function(r) list(file_name = file_label(r$file_path), st = r$st, bank = r$bank)),
                  months, tol)
  })
  # order: customer name, bank, account number
  ord <- order(vapply(accounts, function(a) toupper(squish(a$account_name %||% "")), ""),
               vapply(accounts, `[[`, "", "bank_id"),
               vapply(accounts, function(a) a$account_no %||% "", ""))
  list(months = months, accounts = accounts[ord])
}

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || is.na(a)) b else a
