#' Read a PDF statement into a word/line structure that every bank parser uses
#'
#' @return
#' A "doc" list:
#'   path, file_name, n_pages, text (character, one per page),
#'   words : data.frame(page, line_id, x, x2, y, width, height, text)
#'   lines : data.frame(line_id, page, y, text)
#'
#' @param
#' 1. path     : full path of the PDF
#' 2. password : user password of an encrypted PDF ("" when none)
#'

read_statement_pdf <- function(path, password = "") {

  if (!file.exists(path)) stop_bsp("PDF_NOT_FOUND", paste0("File tidak ditemukan: ", path))

  open_error <- function(e) {
    msg <- conditionMessage(e)
    if (grepl("password|encrypt", msg, ignore.case = TRUE)) {
      stop_bsp("PDF_PASSWORD", if (password == "") "PDF terkunci password - isi kolom Password"
                               else "Password PDF salah")
    }
    stop_bsp("PDF_UNREADABLE", paste0("PDF tidak bisa dibaca: ", msg))
  }

  # pdftools opens PDFs that only carry an owner password without asking
  text  <- tryCatch(pdftools::pdf_text(path, upw = password, opw = password), error = open_error)
  pages <- tryCatch(pdftools::pdf_data(path, upw = password, opw = password), error = open_error)

  if (sum(nchar(trimws(text))) < 20) {
    stop_bsp("PDF_NO_TEXT", "PDF tidak memiliki lapisan teks (kemungkinan hasil scan)")
  }

  words_list <- lapply(seq_along(pages), function(p) {
    w <- pages[[p]]
    if (nrow(w) == 0) return(NULL)
    data.frame(page = p, x = w$x, x2 = w$x + w$width, y = w$y,
               width = w$width, height = w$height, text = w$text,
               stringsAsFactors = FALSE)
  })
  words <- do.call(rbind, words_list)

  words <- assign_lines(words)
  words <- drop_ghost_words(words)

  lines <- build_lines(words)

  list(path = path, file_name = file_label(path), n_pages = length(text),
       text = text, words = words, lines = lines)
}

#' Group words into visual lines (same page, y within tolerance)
assign_lines <- function(words, tol = 2) {
  words <- words[order(words$page, words$y, words$x), ]
  line_id <- integer(nrow(words))
  current <- 0L
  last_page <- -1L
  line_y <- -Inf
  for (i in seq_len(nrow(words))) {
    if (words$page[i] != last_page || words$y[i] - line_y > tol) {
      current <- current + 1L
      line_y <- words$y[i]
      last_page <- words$page[i]
    }
    line_id[i] <- current
  }
  words$line_id <- line_id
  words <- words[order(words$line_id, words$x), ]
  rownames(words) <- NULL
  words
}

#' Some PDFs (e.g. BNI) draw glyphs twice, which pdftools reports as an extra
#' word lying inside another word's box. Remove words fully covered by a wider
#' word on the same line.
drop_ghost_words <- function(words, tol = 1) {
  keep <- rep(TRUE, nrow(words))
  idx_by_line <- split(seq_len(nrow(words)), words$line_id)
  for (idx in idx_by_line) {
    if (length(idx) < 2) next
    for (i in idx) {
      for (j in idx) {
        if (i == j || !keep[j]) next
        if (words$width[j] > words$width[i] &&
            words$x[i] >= words$x[j] - tol && words$x2[i] <= words$x2[j] + tol) {
          keep[i] <- FALSE
          break
        }
      }
    }
  }
  words[keep, ]
}

build_lines <- function(words) {
  sp <- split(words, words$line_id)
  data.frame(line_id = as.integer(names(sp)),
             page = vapply(sp, function(w) w$page[1], integer(1)),
             y = vapply(sp, function(w) min(w$y), numeric(1)),
             x = vapply(sp, function(w) min(w$x), numeric(1)),
             text = vapply(sp, function(w) paste(w$text, collapse = " "), character(1)),
             stringsAsFactors = FALSE)
}

#' Words of one line, sorted by x
line_words <- function(doc, id) {
  w <- doc$words[doc$words$line_id == id, ]
  w[order(w$x), ]
}

#' Text of the words of a line whose left edge lies in [x_from, x_to)
words_between <- function(w, x_from, x_to) {
  sel <- w[w$x >= x_from & w$x < x_to, ]
  paste(sel$text, collapse = " ")
}

#' All text of the document (pages joined) for detection
doc_all_text <- function(doc) paste(doc$text, collapse = "\n")
