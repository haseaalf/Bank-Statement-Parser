#' Result files read back by the Excel macro (all UTF-8 CSV, one line per record,
#' no embedded line breaks so VBA can import them with Workbooks.OpenText)

one_line <- function(x) {
  x <- gsub("[\r\n]+", " | ", as.character(x))
  x[is.na(x)] <- ""
  x
}

write_result_csv <- function(df, file) {
  df[] <- lapply(df, function(col) if (is.character(col)) one_line(col) else col)
  utils::write.csv(df, file, row.names = FALSE, na = "", fileEncoding = "UTF-8")
}

fmt_date <- function(d) ifelse(is.na(d), "", format(d, "%d/%m/%Y"))

#' One row of run_result.csv
result_row <- function(file_path, bank = NULL, st = NULL, ver = NULL, output_file = "",
                       status = "ERROR", message = "", group_id = "") {
  data.frame(file_path = file_path,
             file_name = basename(file_path),
             group_id = group_id,
             bank_id = if (is.null(bank)) "" else bank$id,
             variant = if (is.null(bank)) "" else bank$variant,
             account_no = if (is.null(st)) "" else one_line(st$meta$account_no),
             account_name = if (is.null(st)) "" else one_line(st$meta$account_name),
             currency = if (is.null(st)) "" else one_line(st$meta$currency),
             period_start = if (is.null(st)) "" else fmt_date(st$meta$period_start),
             period_end = if (is.null(st)) "" else fmt_date(st$meta$period_end),
             rows = if (is.null(ver)) NA else ver$summary$rows,
             opening = if (is.null(ver)) NA else ver$summary$opening,
             opening_source = if (is.null(ver)) "" else ver$summary$opening_source,
             total_debit = if (is.null(ver)) NA else ver$summary$total_debit,
             total_credit = if (is.null(ver)) NA else ver$summary$total_credit,
             closing = if (is.null(ver)) NA else ver$summary$closing,
             n_error = if (is.null(ver)) NA else sum(ver$findings$severity == "ERROR"),
             n_warn = if (is.null(ver)) NA else sum(ver$findings$severity == "WARN"),
             status = status,
             output_file = output_file,
             message = message,
             stringsAsFactors = FALSE)
}

#' Rows of verification_detail.csv for one statement
detail_rows <- function(file_name, output_file, st, ver) {
  f <- ver$findings
  if (nrow(f) == 0) return(NULL)
  t <- st$txns
  has_row <- !is.na(f$row) & f$row >= 1 & f$row <= nrow(t)
  data.frame(file_name = file_name,
             output_file = basename(output_file),
             excel_row = ifelse(has_row, f$row + 1, NA),   # +1 for the header row
             date = ifelse(has_row, format(t$date[pmax(1, f$row)], "%d/%m/%Y"), ""),
             description = ifelse(has_row, sub("\n.*$", "", t$description[pmax(1, f$row)]), ""),
             debit = ifelse(has_row, t$debit[pmax(1, f$row)], NA),
             credit = ifelse(has_row, t$credit[pmax(1, f$row)], NA),
             balance = ifelse(has_row, t$balance[pmax(1, f$row)], NA),
             check = f$check, severity = f$severity, message = f$message,
             stringsAsFactors = FALSE)
}
