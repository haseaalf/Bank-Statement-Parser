#' Write the recap workbook in the format of the "rekap mutasi" example:
#'   sheet REKAP   : Bank | Nomor Rekening | Nama Nasabah | Debit | Kredit | Saldo | %
#'                   + Jumlah, "Penjualan L/R Periode ..." (filled in by the user) and
#'                   "% Transaksi Bank" (= total kredit / penjualan)
#'   one sheet per account : BULAN | DEBET | KREDIT | SALDO AKHIR BULAN, + Total and Ratas;
#'                   warnings are written in column E next to the month row

RECAP_BLUE  <- "#00B0F0"
ACCOUNTING  <- '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
MONTH_FMT   <- "[$-409]mmmm\\-yy;@"

#' Unique, valid sheet names (account number; bank prefix when numbers repeat)
recap_sheet_names <- function(accounts) {
  base <- vapply(accounts, function(a) gsub("[\\[\\]:*?/\\\\]", "_", a$account_no %||% "NA", perl = TRUE), "")
  dup <- base %in% base[duplicated(base)]
  base[dup] <- paste(vapply(accounts[dup], `[[`, "", "bank_id"), base[dup])
  base <- substr(base, 1, 31)
  make.unique(base, sep = "_")
}

write_recap_xlsx <- function(recap, file) {
  st <- function(...) openxlsx::createStyle(...)
  wb <- openxlsx::createWorkbook()
  accounts <- recap$accounts
  sheets <- recap_sheet_names(accounts)
  n <- length(accounts)

  # ---- REKAP -------------------------------------------------------------------
  openxlsx::addWorksheet(wb, "REKAP")
  hdr <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "center", valign = "center",
            wrapText = TRUE, fgFill = RECAP_BLUE, border = "TopBottomLeftRight", borderStyle = "medium")
  openxlsx::writeData(wb, "REKAP", t(c("Bank", "Nomor Rekening", "Nama Nasabah", " Debit", "Kredit", "Saldo", "%")),
                      startRow = 1, colNames = FALSE)
  openxlsx::addStyle(wb, "REKAP", hdr, rows = 1, cols = 1:7)

  body_c <- st(fontName = "Tahoma", fontSize = 8, halign = "center", wrapText = TRUE,
               border = "TopBottomLeftRight", borderStyle = "medium")
  body_l <- st(fontName = "Tahoma", fontSize = 8, wrapText = TRUE, border = "TopBottomLeftRight", borderStyle = "medium")
  body_n <- st(fontName = "Tahoma", fontSize = 8, halign = "right", wrapText = TRUE, numFmt = ACCOUNTING,
               border = "TopBottomLeftRight", borderStyle = "medium")
  body_p <- st(fontName = "Tahoma", fontSize = 8, halign = "right", valign = "center", wrapText = TRUE,
               numFmt = "0.00%", border = "TopBottomLeftRight", borderStyle = "medium")
  warn_st <- st(fontName = "Tahoma", fontSize = 8, fontColour = "#C00000")

  jr <- n + 2                                   # "Jumlah" row
  for (i in seq_len(n)) {
    a <- accounts[[i]]; r <- i + 1
    q <- paste0("'", sheets[i], "'!")
    tr <- length(recap$months) + 2              # Total row on the account sheet
    openxlsx::writeData(wb, "REKAP", t(c(a$bank_id, a$account_no %||% "", a$account_name %||% "")),
                        startRow = r, colNames = FALSE)
    openxlsx::writeFormula(wb, "REKAP", paste0(q, "B", tr), startCol = 4, startRow = r)
    openxlsx::writeFormula(wb, "REKAP", paste0(q, "C", tr), startCol = 5, startRow = r)
    openxlsx::writeFormula(wb, "REKAP", paste0(q, "D", tr + 1), startCol = 6, startRow = r)
    openxlsx::writeFormula(wb, "REKAP", paste0("IFERROR(F", r, "/F$", jr, ",0)"), startCol = 7, startRow = r)
    openxlsx::addStyle(wb, "REKAP", body_c, rows = r, cols = 1:2)
    openxlsx::addStyle(wb, "REKAP", body_l, rows = r, cols = 3)
    openxlsx::addStyle(wb, "REKAP", body_n, rows = r, cols = 4:6, gridExpand = TRUE)
    openxlsx::addStyle(wb, "REKAP", body_p, rows = r, cols = 7)
    nw <- sum(nzchar(a$table$warning))
    if (nw > 0) {
      openxlsx::writeData(wb, "REKAP", paste0(nw, " bulan perlu dicek - lihat kolom E di sheet ", sheets[i]),
                          startCol = 8, startRow = r)
      openxlsx::addStyle(wb, "REKAP", warn_st, rows = r, cols = 8)
    }
  }

  # Jumlah / Penjualan / % Transaksi Bank
  tot_l <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", valign = "top", wrapText = TRUE,
              fgFill = RECAP_BLUE, border = "TopBottomLeftRight", borderStyle = "medium")
  tot_n <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "right", valign = "top",
              wrapText = TRUE, numFmt = ACCOUNTING, fgFill = RECAP_BLUE, border = "TopBottomLeftRight",
              borderStyle = "medium")
  tot_p <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "right", valign = "top",
              wrapText = TRUE, numFmt = "0.00%", fgFill = RECAP_BLUE, border = "TopBottomLeftRight",
              borderStyle = "medium")
  lab <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "left", valign = "center",
            wrapText = TRUE, fgFill = RECAP_BLUE, border = "TopBottomLeftRight", borderStyle = "medium")
  val_n <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "center", valign = "center",
              wrapText = TRUE, numFmt = ACCOUNTING, fgFill = RECAP_BLUE, border = "TopBottomLeftRight",
              borderStyle = "medium")
  val_p <- st(fontName = "Tahoma", fontSize = 8, textDecoration = "bold", halign = "center", valign = "center",
              wrapText = TRUE, numFmt = "0.00%", fgFill = RECAP_BLUE, border = "TopBottomLeftRight",
              borderStyle = "medium")

  openxlsx::writeData(wb, "REKAP", "Jumlah", startCol = 1, startRow = jr)
  if (n > 0) for (k in 4:7) {
    col <- LETTERS[k]
    openxlsx::writeFormula(wb, "REKAP", paste0("SUM(", col, "2:", col, n + 1, ")"), startCol = k, startRow = jr)
  }
  openxlsx::mergeCells(wb, "REKAP", cols = 1:3, rows = jr)
  openxlsx::addStyle(wb, "REKAP", tot_l, rows = jr, cols = 1:3)
  openxlsx::addStyle(wb, "REKAP", tot_n, rows = jr, cols = 4:6)
  openxlsx::addStyle(wb, "REKAP", tot_p, rows = jr, cols = 7)

  openxlsx::writeData(wb, "REKAP", "Penjualan L/R Periode ...", startCol = 1, startRow = jr + 1)
  openxlsx::mergeCells(wb, "REKAP", cols = 1:3, rows = jr + 1)
  openxlsx::addStyle(wb, "REKAP", lab, rows = jr + 1, cols = 1:3)
  openxlsx::addStyle(wb, "REKAP", val_n, rows = jr + 1, cols = 4:7)

  openxlsx::writeData(wb, "REKAP", "% Transaksi Bank", startCol = 1, startRow = jr + 2)
  openxlsx::writeFormula(wb, "REKAP", paste0('IFERROR(E', jr, '/E', jr + 1, ',"")'), startCol = 5, startRow = jr + 2)
  openxlsx::mergeCells(wb, "REKAP", cols = 1:3, rows = jr + 2)
  openxlsx::addStyle(wb, "REKAP", lab, rows = jr + 2, cols = 1:3)
  openxlsx::addStyle(wb, "REKAP", val_p, rows = jr + 2, cols = 4:7)

  # widths of the example; amount columns grow when the numbers are longer (avoids "#####")
  fit <- function(base, x) {
    x <- x[is.finite(x)]
    if (length(x) == 0) return(base)
    max(base, 0.8 * max(nchar(formatC(round(x), format = "f", digits = 0, big.mark = ","))) + 1.5)
  }
  dsum <- vapply(accounts, function(a) sum(a$table$debit, na.rm = TRUE), 0)
  csum <- vapply(accounts, function(a) sum(a$table$credit, na.rm = TRUE), 0)
  savg <- vapply(accounts, function(a) mean(a$table$saldo, na.rm = TRUE), 0)
  openxlsx::setColWidths(wb, "REKAP", cols = 1:7,
                         widths = c(4.21875, 15, 14, fit(11.88671875, c(dsum, sum(dsum))),
                                    fit(11.88671875, c(csum, sum(csum))), fit(10.5546875, c(savg, sum(savg))),
                                    7.21875) - 0.71)
  openxlsx::setRowHeights(wb, "REKAP", rows = 1, heights = 21)
  openxlsx::setRowHeights(wb, "REKAP", rows = 2:(jr + 2), heights = 15)

  # ---- one sheet per account ---------------------------------------------------
  months <- month_first(recap$months)
  for (i in seq_len(n)) {
    a <- accounts[[i]]; s <- sheets[i]; tb <- a$table
    fill <- RECAP_BLUE                          # same blue header on every account sheet
    openxlsx::addWorksheet(wb, s)
    openxlsx::writeData(wb, s, t(c("BULAN", "DEBET", "KREDIT", "SALDO AKHIR BULAN")), startRow = 1, colNames = FALSE)
    openxlsx::addStyle(wb, s, st(fontName = "Calibri", fontSize = 11, textDecoration = "bold", halign = "center",
                                 fgFill = fill, border = "TopBottomLeftRight", numFmt = MONTH_FMT),
                       rows = 1, cols = 1)
    openxlsx::addStyle(wb, s, st(fontName = "Calibri", fontSize = 11, textDecoration = "bold", halign = "center",
                                 fgFill = fill, border = "TopBottomLeftRight", numFmt = "#,##0"),
                       rows = 1, cols = 2:4)

    m_st <- st(fontName = "Tahoma", fontSize = 9, halign = "left", numFmt = MONTH_FMT, border = "TopBottomLeftRight")
    v_st <- st(fontName = "Tahoma", fontSize = 8, valign = "center", wrapText = TRUE, numFmt = "#,##0.00",
               border = "TopBottomLeftRight")
    z_st <- st(fontName = "Tahoma", fontSize = 9, numFmt = "#,##0", border = "TopBottomLeftRight")
    for (k in seq_along(months)) {
      r <- k + 1
      openxlsx::writeData(wb, s, months[k], startCol = 1, startRow = r)
      openxlsx::addStyle(wb, s, m_st, rows = r, cols = 1)
      if (tb$has_data[k]) {
        vals <- c(tb$debit[k], tb$credit[k], tb$saldo[k])
        openxlsx::writeData(wb, s, t(vals), startCol = 2, startRow = r, colNames = FALSE)
        # like the example: amounts with decimals style, zero amounts plain
        for (cc in 1:3) openxlsx::addStyle(wb, s, if (vals[cc] == 0) z_st else v_st, rows = r, cols = cc + 1)
      } else {
        openxlsx::addStyle(wb, s, z_st, rows = r, cols = 2:4)      # month without statement: left empty
      }
      if (nzchar(tb$warning[k])) {
        openxlsx::writeData(wb, s, paste("PERINGATAN:", tb$warning[k]), startCol = 5, startRow = r)
        openxlsx::addStyle(wb, s, warn_st, rows = r, cols = 5)
      }
    }
    tr <- length(months) + 2
    last <- tr - 1
    t_st <- st(fontName = "Tahoma", fontSize = 9, textDecoration = "bold", halign = "center", fgFill = fill,
               border = "TopBottomLeftRight", numFmt = "#,##0")
    openxlsx::writeData(wb, s, "Total", startCol = 1, startRow = tr)
    openxlsx::writeFormula(wb, s, paste0("SUM(B2:B", last, ")"), startCol = 2, startRow = tr)
    openxlsx::writeFormula(wb, s, paste0("SUM(C2:C", last, ")"), startCol = 3, startRow = tr)
    openxlsx::writeData(wb, s, "Ratas", startCol = 1, startRow = tr + 1)
    for (cc in 2:4) {
      col <- LETTERS[cc]
      openxlsx::writeFormula(wb, s, paste0("AVERAGE(", col, "2:", col, last, ")"), startCol = cc, startRow = tr + 1)
    }
    openxlsx::addStyle(wb, s, t_st, rows = tr:(tr + 1), cols = 1:4, gridExpand = TRUE)
    openxlsx::addStyle(wb, s, st(fontName = "Tahoma", fontSize = 9, textDecoration = "bold", halign = "center",
                                 fgFill = fill, border = "TopBottomLeftRight", numFmt = MONTH_FMT),
                       rows = tr:(tr + 1), cols = 1)
    openxlsx::setColWidths(wb, s, cols = 1:4, widths = c(11, 13.88671875, 13.88671875, 18.88671875) - 0.71)
  }

  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  clean_xlsx_package(file)
  invisible(file)
}
