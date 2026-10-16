#' Write transactions in the standard output format (identical to the sample
#' "input output sample/*.xlsx"): one sheet named after the bank, columns A-E.

OUTPUT_HEADERS <- c("Transaction Date\n(DD/MM)", "Description", "Debit", "Credit", "Balance")
OUTPUT_WIDTHS  <- c(17, 51, 34, 34, 34)

#' @param txns       data.frame(date, description, debit, credit, balance)
#' @param sheet_name usually the bank id (e.g. "BRI")
#' @param file       full path of the .xlsx to create
write_output_xlsx <- function(txns, sheet_name, file) {
  wb <- openxlsx::createWorkbook()
  sheet_name <- substr(gsub("[\\[\\]:*?/\\\\]", "_", sheet_name, perl = TRUE), 1, 31)
  openxlsx::addWorksheet(wb, sheet_name)

  out <- data.frame(date = format(txns$date, "%d/%m"),
                    description = txns$description,
                    debit = round(txns$debit, 2),
                    credit = round(txns$credit, 2),
                    balance = round(txns$balance, 2),
                    stringsAsFactors = FALSE)
  names(out) <- OUTPUT_HEADERS

  openxlsx::writeData(wb, sheet_name, out, startRow = 1, startCol = 1,
                      colNames = TRUE, keepNA = FALSE)

  hdr <- openxlsx::createStyle(fontName = "Arial", fontSize = 11,
                               halign = "center", valign = "center")
  hdr_wrap <- openxlsx::createStyle(fontName = "Arial", fontSize = 11,
                                    halign = "center", valign = "center", wrapText = TRUE)
  openxlsx::addStyle(wb, sheet_name, hdr, rows = 1, cols = 2:5)
  openxlsx::addStyle(wb, sheet_name, hdr_wrap, rows = 1, cols = 1)

  if (nrow(out) > 0) {
    num <- openxlsx::createStyle(numFmt = "0.00")
    openxlsx::addStyle(wb, sheet_name, num, rows = 1 + seq_len(nrow(out)), cols = 3:5,
                       gridExpand = TRUE)
  }

  # openxlsx adds 0.71 padding to every width; compensate so the file stores 17/51/34
  openxlsx::setColWidths(wb, sheet_name, cols = 1:5, widths = OUTPUT_WIDTHS - 0.71)
  openxlsx::setRowHeights(wb, sheet_name, rows = 1, heights = 34)

  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  clean_xlsx_package(file)
  invisible(file)
}

#' openxlsx 4.2.5 declares drawing parts it does not write (relationships and
#' content-type overrides pointing to missing files). Excel tolerates this but
#' strict readers (openpyxl, some converters) fail, so remove the dangling entries.
clean_xlsx_package <- function(file) {
  tmp <- tempfile("bsp_xlsx_")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  utils::unzip(file, exdir = tmp)
  parts <- list.files(tmp, recursive = TRUE, all.files = TRUE)
  changed <- FALSE

  for (rel in parts[grepl("_rels/[^/]+\\.rels$", parts)]) {
    path <- file.path(tmp, rel)
    xml <- paste(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = "")
    base <- dirname(dirname(rel))                       # folder the .rels describes
    items <- regmatches(xml, gregexpr("<Relationship [^>]*/>", xml))[[1]]
    for (it in items) {
      if (grepl('TargetMode="External"', it)) next
      target <- sub('.*Target="([^"]+)".*', "\\1", it)
      full <- if (startsWith(target, "/")) sub("^/", "", target) else file.path(base, target)
      full <- normalizePath(file.path(tmp, full), winslash = "/", mustWork = FALSE)
      if (!file.exists(full)) { xml <- sub(it, "", xml, fixed = TRUE); changed <- TRUE }
    }
    writeLines(xml, path, useBytes = TRUE)
  }

  ct_path <- file.path(tmp, "[Content_Types].xml")
  ct <- paste(readLines(ct_path, warn = FALSE, encoding = "UTF-8"), collapse = "")
  for (it in regmatches(ct, gregexpr("<Override [^>]*/>", ct))[[1]]) {
    part <- sub('.*PartName="/([^"]+)".*', "\\1", it)
    if (!file.exists(file.path(tmp, part))) { ct <- sub(it, "", ct, fixed = TRUE); changed <- TRUE }
  }
  writeLines(ct, ct_path, useBytes = TRUE)

  if (changed) {
    # Excel rejects packages with folder entries: zip the files only, [Content_Types].xml first
    files <- c("[Content_Types].xml", setdiff(list.files(tmp, recursive = TRUE, all.files = TRUE),
                                              "[Content_Types].xml"))
    unlink(file)
    zip::zip(file, files = files, root = tmp, mode = "mirror", include_directories = FALSE)
  }
  invisible(file)
}

#' Resolve the output file name according to the collision policy
#' @param on_exists "overwrite" or "suffix" (adds " (2)", " (3)", ...)
output_path <- function(dir, base_name, on_exists = "suffix") {
  base_name <- gsub("[<>:\"/\\\\|?*]", "_", base_name)
  win <- function(f) normalizePath(f, winslash = "\\", mustWork = FALSE)
  f <- file.path(dir, paste0(base_name, ".xlsx"))
  if (on_exists == "overwrite" || !file.exists(f)) return(win(f))
  k <- 2
  repeat {
    f <- file.path(dir, paste0(base_name, " (", k, ").xlsx"))
    if (!file.exists(f)) return(win(f))
    k <- k + 1
  }
}
