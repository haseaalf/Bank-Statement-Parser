#' Layout helpers for table-like statements

#' Lines of the transaction table: on every page, the lines after the first
#' line matching `header` and before the first line matching `stop`.
#' Pages without the header are skipped unless `header_optional` (then the
#' whole page is used, still stopping at `stop`).
table_lines <- function(doc, header, stop = NULL, skip = NULL, header_optional = FALSE,
                        compact_header = FALSE) {
  L <- doc$lines
  out <- list()
  for (p in seq_len(doc$n_pages)) {
    pl <- L[L$page == p, ]
    # compact_header: match the header with spaces removed (letter-spaced fonts, e.g. BCA)
    h <- which(grepl(header, if (compact_header) gsub(" ", "", pl$text) else pl$text))
    if (length(h) == 0) {
      if (!header_optional) next
      start <- 1
    } else start <- h[1] + 1
    if (start > nrow(pl)) next
    body <- pl[start:nrow(pl), ]
    if (!is.null(stop)) {
      s <- which(grepl(stop, body$text))
      if (length(s) > 0) body <- body[seq_len(s[1] - 1), ]
    }
    if (!is.null(skip) && nrow(body) > 0) body <- body[!grepl(skip, body$text), ]
    out[[length(out) + 1]] <- body
  }
  if (length(out) == 0) return(L[0, ])
  do.call(rbind, out)
}

#' Assign every line to the nearest anchor line (same page) by vertical
#' distance. Used for layouts where a transaction's text is vertically
#' centred around the line holding its amounts (e.g. Mandiri, Maybank).
#' @return integer vector: index into `anchors` for each line (NA if none)
nearest_anchor <- function(lines, anchors) {
  vapply(seq_len(nrow(lines)), function(i) {
    cand <- which(anchors$page == lines$page[i])
    if (length(cand) == 0) return(NA_integer_)
    cand[which.min(abs(anchors$y[cand] - lines$y[i]))]
  }, integer(1))
}

#' Horizontal table rules of a page (y positions in PDF points).
#'
#' pdftools does not expose vector graphics, so the page is rendered and pixel
#' rows that are uniformly non-white across [x_from, x_to] are taken as rules
#' (thin runs) or as the edges of shaded bands. Used by layouts whose
#' multi-line cells cannot be grouped from the text alone (e.g. BNI).
table_row_rules <- function(doc, page, x_from, x_to, dpi = 144, min_cover = 0.95) {
  img <- tryCatch(suppressMessages(pdftools::pdf_render_page(doc$path, page = page, dpi = dpi,
                                                             numeric = FALSE)),
                  error = function(e) NULL)
  if (is.null(img)) return(numeric())
  k <- dpi / 72
  g <- as.integer(img[1, , ]); dim(g) <- dim(img)[2:3]          # width x height
  x0 <- max(1, round(x_from * k)); x1 <- min(nrow(g), round(x_to * k))
  cover <- colMeans(g[x0:x1, , drop = FALSE] < 250)
  hit <- which(cover >= min_cover)
  if (length(hit) == 0) return(numeric())
  runs <- split(hit, cumsum(c(1, diff(hit) > 1)))
  edges <- unlist(lapply(runs, function(r) if (length(r) <= 4) mean(r) else c(min(r), max(r))))
  edges <- sort(edges / k)
  # merge double lines (adjacent cell borders)
  edges[c(TRUE, diff(edges) > 3)]
}

#' Split the word table of a line into column texts given left boundaries.
#' @param bounds named numeric vector of column start x (sorted)
split_columns <- function(w, bounds) {
  edges <- c(bounds, Inf)
  res <- vapply(seq_along(bounds), function(k) {
    paste(w$text[w$x >= edges[k] & w$x < edges[k + 1]], collapse = " ")
  }, character(1))
  names(res) <- names(bounds)
  res
}

#' Text of a document zone as one string (for meta regexes): lines joined by \n
lines_text <- function(lines) paste(lines$text, collapse = "\n")

#' Build a txns data.frame from a list of row lists (date, desc, debit, credit, balance, page)
rows_to_txns <- function(rows) {
  if (length(rows) == 0) return(new_txns())
  num <- function(field) vapply(rows, function(r) { v <- r[[field]]; if (is.null(v) || length(v) == 0) NA_real_ else as.numeric(v) }, 0)
  data.frame(date = as.Date(num("date"), origin = "1970-01-01"),
             description = vapply(rows, function(r) join_desc(r$desc), ""),
             debit = num("debit"), credit = num("credit"), balance = num("balance"),
             page = as.integer(num("page")),
             stringsAsFactors = FALSE)
}
