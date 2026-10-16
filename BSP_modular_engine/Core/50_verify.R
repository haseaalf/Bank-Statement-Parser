#' Statement verification
#'
#' Checks never block the output; they produce findings with a severity:
#'   ERROR : the numbers do not add up (running balance, totals, opening/closing)
#'   WARN  : something to review (derived opening balance, date outside period...)
#'   INFO  : how a figure was obtained

new_findings <- function() {
  data.frame(row = integer(), check = character(), severity = character(),
             message = character(), stringsAsFactors = FALSE)
}

add_finding <- function(f, row, check, severity, message) {
  rbind(f, data.frame(row = as.integer(row), check = check, severity = severity,
                      message = message, stringsAsFactors = FALSE))
}

fmt_num <- function(x) formatC(x, format = "f", digits = 2, big.mark = ",")

#' Number of rows whose balance does not follow from the previous row
chain_breaks <- function(t, tol = 0.01) {
  if (nrow(t) < 2) return(0)
  ok <- !is.na(t$balance)
  sum(vapply(2:nrow(t), function(i) {
    if (!ok[i] || !ok[i - 1]) return(FALSE)
    abs(t$balance[i - 1] - t$debit[i] + t$credit[i] - t$balance[i]) > tol
  }, logical(1)))
}

#' Some banks list the newest transaction first (e.g. CIMB). Return the rows in
#' chronological order: keep the PDF order or reverse it, whichever makes the
#' running balance consistent.
fix_order_by_chain <- function(t) {
  if (nrow(t) < 2) return(t)
  r <- t[rev(seq_len(nrow(t))), ]
  keep <- chain_breaks(t) <= chain_breaks(r) &&
    !(chain_breaks(t) == chain_breaks(r) && t$date[1] > t$date[nrow(t)])
  out <- if (keep) t else r
  rownames(out) <- NULL
  out
}

#' Opening balance: stated in the PDF, else derived from the first row's balance
opening_balance <- function(st) {
  if (!is.na(st$stated$opening)) return(list(value = st$stated$opening, source = "stated"))
  t <- st$txns
  first_bal <- which(!is.na(t$balance))[1]
  if (is.na(first_bal)) return(list(value = NA_real_, source = "missing"))
  v <- t$balance[first_bal] - sum(t$credit[seq_len(first_bal)]) + sum(t$debit[seq_len(first_bal)])
  list(value = round(v, 2), source = "derived")
}

#' Fill balances the PDF does not print (e.g. BCA prints only end-of-day
#' balances) from the running balance. Returns the statement with
#' txns$balance_source = "stated" / "derived".
complete_balances <- function(st) {
  t <- st$txns
  t$balance_source <- ifelse(is.na(t$balance), "derived", "stated")
  ob <- opening_balance(st)
  prev <- ob$value
  for (i in seq_len(nrow(t))) {
    if (is.na(t$balance[i]) && !is.na(prev)) t$balance[i] <- round(prev - t$debit[i] + t$credit[i], 2)
    prev <- t$balance[i]
  }
  st$txns <- t
  st$opening <- ob
  st
}

#' Verify one statement.
#' @param tol tolerance for amount comparisons
#' @return list(findings, status, summary)
verify_statement <- function(st, tol = 0.01) {
  f <- new_findings()
  t <- st$txns
  n <- nrow(t)
  ob <- if (!is.null(st$opening)) st$opening else opening_balance(st)

  if (n == 0) {
    f <- add_finding(f, NA, "rows", "ERROR", "Tidak ada transaksi yang terbaca")
  }

  if (ob$source == "derived") {
    f <- add_finding(f, NA, "opening", "WARN",
                     paste0("Saldo awal tidak tercantum di PDF; dihitung dari baris pertama = ", fmt_num(ob$value)))
  }

  # 1. Row level --------------------------------------------------------------
  prev <- ob$value
  for (i in seq_len(n)) {
    if (t$debit[i] < 0 || t$credit[i] < 0) {
      f <- add_finding(f, i, "amount", "ERROR", "Nilai debit/kredit negatif")
    }
    if (t$debit[i] > 0 && t$credit[i] > 0) {
      f <- add_finding(f, i, "amount", "WARN", "Debit dan kredit terisi pada baris yang sama")
    }
    if (is.na(t$date[i])) {
      f <- add_finding(f, i, "date", "ERROR", "Tanggal tidak terbaca")
    }
    if (is.na(t$balance[i])) {
      f <- add_finding(f, i, "balance", "ERROR", "Saldo tidak terbaca")
    } else if (!is.na(prev)) {
      expected <- round(prev - t$debit[i] + t$credit[i], 2)
      if (abs(expected - t$balance[i]) > tol) {
        f <- add_finding(f, i, "running_balance", "ERROR",
                         paste0("Saldo tidak sesuai: seharusnya ", fmt_num(expected),
                                ", tertulis ", fmt_num(t$balance[i]),
                                " (selisih ", fmt_num(t$balance[i] - expected), ")"))
      }
    }
    prev <- t$balance[i]
  }

  # 2. Dates ------------------------------------------------------------------
  ps <- st$meta$period_start
  pe <- st$meta$period_end
  if (n > 0 && !is.na(ps) && !is.na(pe)) {
    out <- which(!is.na(t$date) & (t$date < ps | t$date > pe))
    for (i in out) f <- add_finding(f, i, "date", "WARN",
                                    paste0("Tanggal ", format(t$date[i], "%d/%m/%Y"), " di luar periode statement"))
  }
  if (n > 1) {
    back <- which(diff(as.numeric(t$date)) < 0) + 1
    for (i in back) f <- add_finding(f, i, "date", "WARN", "Tanggal lebih awal dari baris sebelumnya")
  }

  # 3. Statement level --------------------------------------------------------
  s <- st$stated
  last_bal <- if (n > 0) t$balance[n] else ob$value
  sum_dr <- round(sum(t$debit), 2)
  sum_cr <- round(sum(t$credit), 2)
  n_dr <- sum(t$debit > 0)
  n_cr <- sum(t$credit > 0)

  if (!is.na(s$closing)) {
    if (is.na(last_bal) || abs(last_bal - s$closing) > tol) {
      f <- add_finding(f, n, "closing", "ERROR",
                       paste0("Saldo akhir PDF ", fmt_num(s$closing), " berbeda dengan saldo baris terakhir ", fmt_num(last_bal)))
    }
  } else {
    f <- add_finding(f, NA, "closing", "INFO", "Saldo akhir tidak tercantum di PDF")
  }
  if (!is.na(ob$value) && !is.na(last_bal) && abs(ob$value - sum_dr + sum_cr - last_bal) > tol) {
    f <- add_finding(f, NA, "balance_equation", "ERROR",
                     "Saldo awal - total debit + total kredit tidak sama dengan saldo akhir")
  }
  if (!is.na(s$total_debit) && abs(s$total_debit - sum_dr) > tol) {
    f <- add_finding(f, NA, "total_debit", "ERROR",
                     paste0("Total debit PDF ", fmt_num(s$total_debit), " vs hasil parse ", fmt_num(sum_dr)))
  }
  if (!is.na(s$total_credit) && abs(s$total_credit - sum_cr) > tol) {
    f <- add_finding(f, NA, "total_credit", "ERROR",
                     paste0("Total kredit PDF ", fmt_num(s$total_credit), " vs hasil parse ", fmt_num(sum_cr)))
  }
  if (!is.na(s$n_debit) && s$n_debit != n_dr) {
    f <- add_finding(f, NA, "count_debit", "ERROR",
                     paste0("Jumlah transaksi debit PDF ", s$n_debit, " vs hasil parse ", n_dr))
  }
  if (!is.na(s$n_credit) && s$n_credit != n_cr) {
    f <- add_finding(f, NA, "count_credit", "ERROR",
                     paste0("Jumlah transaksi kredit PDF ", s$n_credit, " vs hasil parse ", n_cr))
  }

  status <- if (any(f$severity == "ERROR")) "ERROR" else if (any(f$severity == "WARN")) "WARN" else "OK"
  list(findings = f, status = status,
       summary = list(rows = n, opening = ob$value, opening_source = ob$source,
                      closing = last_bal, total_debit = sum_dr, total_credit = sum_cr,
                      n_debit = n_dr, n_credit = n_cr))
}
