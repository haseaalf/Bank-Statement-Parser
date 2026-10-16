#' Shared helpers for bank parsers: errors, amounts, dates, transaction table

# ---- Errors -------------------------------------------------------------------

#' Raise an engine error with a machine readable code (shown to the user by VBA)
stop_bsp <- function(code, message) {
  cond <- structure(class = c("bsp_error", "error", "condition"),
                    list(message = message, code = code, call = NULL))
  stop(cond)
}

# ---- Amounts ------------------------------------------------------------------

AMOUNT_US_RE <- "^-?\\d{1,3}(,\\d{3})*(\\.\\d+)?$|^-?\\d+(\\.\\d+)?$"
AMOUNT_ID_RE <- "^-?\\d{1,3}(\\.\\d{3})*(,\\d+)?$|^-?\\d+(,\\d+)?$"

#' Parse amount strings to numeric.
#' style = "us" : 1,234,567.89   style = "id" : 1.234.567,89
#' Trailing markers like DB, CR, K, D are removed (use amount_marker() to read them).
parse_amount <- function(x, style = "us") {
  x <- trimws(as.character(x))
  x <- sub("(DB|CR|DR|K|D|C)$", "", x)
  x <- trimws(x)
  x[x %in% c("", "-")] <- NA
  if (style == "id") {
    x <- gsub(".", "", x, fixed = TRUE)
    x <- sub(",", ".", x, fixed = TRUE)
  } else {
    x <- gsub(",", "", x, fixed = TRUE)
  }
  suppressWarnings(as.numeric(x))
}

#' TRUE when the token looks like an amount in the given style
is_amount <- function(x, style = "us", markers = FALSE) {
  x <- trimws(x)
  if (markers) x <- sub("(DB|CR|DR|K|D)$", "", x)
  grepl(if (style == "id") AMOUNT_ID_RE else AMOUNT_US_RE, x) & grepl("\\d", x)
}

#' Read the debit/credit marker glued to an amount ("10,000.00DB" -> "DB")
amount_marker <- function(x) {
  out <- rep("", length(x))
  hit <- regexpr("(DB|CR|DR|K|D)$", x)
  out[hit > 0] <- regmatches(x, hit)
  out
}

# ---- Dates --------------------------------------------------------------------

MONTHS <- c(jan = 1, feb = 2, mar = 3, apr = 4, may = 5, mei = 5, jun = 6, jul = 7,
            aug = 8, agu = 8, agt = 8, ags = 8, sep = 9, oct = 10, okt = 10,
            nov = 11, dec = 12, des = 12,
            januari = 1, februari = 2, maret = 3, april = 4, juni = 6, juli = 7,
            agustus = 8, september = 9, oktober = 10, november = 11, desember = 12,
            january = 1, february = 2, march = 3, june = 6, july = 7, august = 8,
            october = 10, december = 12)

month_number <- function(x) unname(MONTHS[tolower(x)])

#' Expand a 2 digit year
full_year <- function(y) {
  y <- as.integer(y)
  ifelse(y < 100, 2000L + y, y)
}

make_date <- function(y, m, d) {
  as.Date(sprintf("%04d-%02d-%02d", as.integer(full_year(y)), as.integer(m), as.integer(d)))
}

#' Parse "dd/mm/yy", "dd/mm/yyyy", "dd-mm-yyyy" (order = "dmy") or "mm/dd/yy" ("mdy")
parse_num_date <- function(x, order = "dmy") {
  p <- regmatches(x, regexec("(\\d{1,2})[/.-](\\d{1,2})[/.-](\\d{2,4})", x))
  vapply(p, function(v) {
    if (length(v) < 4) return(NA_real_)
    if (order == "mdy") as.numeric(make_date(v[4], v[2], v[3]))
    else as.numeric(make_date(v[4], v[3], v[2]))
  }, numeric(1)) |> as.Date(origin = "1970-01-01")
}

#' Parse "13 Mar 2026", "01-Feb-26", "01 Agustus 2026"
parse_text_date <- function(x) {
  p <- regmatches(x, regexec("(\\d{1,2})[ -]([A-Za-z]{3,9})[ -](\\d{2,4})", x))
  vapply(p, function(v) {
    if (length(v) < 4) return(NA_real_)
    m <- month_number(v[3])
    if (is.na(m)) return(NA_real_)
    as.numeric(make_date(v[4], m, v[2]))
  }, numeric(1)) |> as.Date(origin = "1970-01-01")
}

#' Year of a "dd/mm" date inside a statement period (handles Dec -> Jan periods)
date_from_day_month <- function(dm, period_start, period_end) {
  p <- regmatches(dm, regexec("(\\d{1,2})/(\\d{1,2})", dm))
  y0 <- as.integer(format(period_start, "%Y"))
  vapply(p, function(v) {
    if (length(v) < 3) return(NA_real_)
    d <- make_date(y0, v[3], v[2])
    if (!is.na(period_end) && d < period_start - 31) d <- make_date(y0 + 1, v[3], v[2])
    as.numeric(d)
  }, numeric(1)) |> as.Date(origin = "1970-01-01")
}

# ---- Text ---------------------------------------------------------------------

squish <- function(x) gsub("\\s+", " ", trimws(x))

#' File name part of a path. Unlike basename(), it has no length limit
#' (basename() fails with "path too long" above 260 characters on Windows).
file_label <- function(x) sub("^.*[\\\\/]", "", x)

#' Join description fragments into a multi-line description (as in the output format)
join_desc <- function(parts) {
  parts <- squish(parts)
  parts <- parts[nchar(parts) > 0]
  paste(parts, collapse = "\n")
}

#' First regex capture group found in text, or NA
grab <- function(text, pattern, group = 1, ignore.case = FALSE) {
  m <- regmatches(text, regexec(pattern, text, ignore.case = ignore.case, perl = TRUE))[[1]]
  if (length(m) > group) m[group + 1] else NA_character_
}

# ---- Standard structures ------------------------------------------------------

#' Empty transaction table in the engine's standard schema
new_txns <- function() {
  data.frame(date = as.Date(character()), description = character(),
             debit = numeric(), credit = numeric(), balance = numeric(),
             page = integer(), stringsAsFactors = FALSE)
}

#' Build a statement object returned by every bank parser
#'
#' @param txns    data.frame(date, description, debit, credit, balance[, page]);
#'                balance may be NA where the PDF does not print it
#' @param meta    list(account_no, account_name, currency, period_start, period_end)
#' @param stated  list(opening, closing, total_debit, total_credit, n_debit, n_credit)
#'                (NA where the PDF does not show the figure)
new_statement <- function(txns, meta = list(), stated = list()) {
  meta_default <- list(account_no = NA_character_, account_name = NA_character_,
                       currency = "IDR", period_start = as.Date(NA), period_end = as.Date(NA))
  stated_default <- list(opening = NA_real_, closing = NA_real_,
                         total_debit = NA_real_, total_credit = NA_real_,
                         n_debit = NA_integer_, n_credit = NA_integer_)
  txns$debit[is.na(txns$debit)] <- 0
  txns$credit[is.na(txns$credit)] <- 0
  if (!"page" %in% names(txns)) txns$page <- NA_integer_
  rownames(txns) <- NULL
  list(txns = txns,
       meta = modifyList(meta_default, meta),
       stated = modifyList(stated_default, stated))
}
