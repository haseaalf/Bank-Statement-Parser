#' Auto-detect which bank / format a statement belongs to

DETECT_MIN_SCORE <- 0.5

#' Score helper for detect() functions: share of `patterns` found in the text
#' (regular expressions, case-insensitive). `required` patterns must all match,
#' otherwise the score is 0.
keyword_score <- function(text, patterns, required = character()) {
  text <- gsub("\\s+", " ", text)   # layout text uses runs of spaces between words
  for (r in required) if (!grepl(r, text, ignore.case = TRUE)) return(0)
  if (length(patterns) == 0) return(1)
  mean(vapply(patterns, function(p) grepl(p, text, ignore.case = TRUE), logical(1)))
}

#' Run every registered detector on the document.
#' @return data.frame(bank_id, variant, score) sorted by score (best first)
score_banks <- function(doc) {
  b <- registered_banks()
  scores <- vapply(b, function(x) {
    s <- tryCatch(x$detect(doc), error = function(e) 0)
    if (!is.numeric(s) || length(s) != 1 || is.na(s)) 0 else s
  }, numeric(1))
  res <- data.frame(bank_id = vapply(b, `[[`, "", "id"),
                    variant = vapply(b, `[[`, "", "variant"),
                    score = scores, stringsAsFactors = FALSE)
  res[order(-res$score), ]
}

#' Pick the bank for a document.
#' @param bank_id,variant user choice; "AUTO"/"" means detect
#' @return list(bank = registry entry or NULL, score, status, message)
resolve_bank <- function(doc, bank_id = "AUTO", variant = "") {
  auto <- is.na(bank_id) || bank_id == "" || toupper(bank_id) == "AUTO"

  if (!auto) {
    cand <- get_bank(bank_id, variant)
    if (!is.null(cand$id)) return(list(bank = cand, score = NA_real_, status = "MANUAL", message = ""))
    # bank chosen but not the variant: detect among its variants
    scores <- vapply(cand, function(x) tryCatch(x$detect(doc), error = function(e) 0), numeric(1))
    best <- cand[[which.max(scores)]]
    return(list(bank = best, score = max(scores), status = "MANUAL",
                message = if (max(scores) < DETECT_MIN_SCORE)
                  "Format PDF kurang cocok dengan bank yang dipilih" else ""))
  }

  s <- score_banks(doc)
  if (nrow(s) == 0 || s$score[1] < DETECT_MIN_SCORE) {
    return(list(bank = NULL, score = if (nrow(s)) s$score[1] else 0, status = "UNKNOWN",
                message = "Bank tidak dikenali, pilih bank secara manual"))
  }
  msg <- ""
  if (nrow(s) > 1 && s$score[2] >= s$score[1] - 0.05 && s$bank_id[2] != s$bank_id[1]) {
    msg <- paste0("Deteksi ragu antara ", s$bank_id[1], " dan ", s$bank_id[2])
  }
  list(bank = get_bank(s$bank_id[1], s$variant[1]), score = s$score[1],
       status = if (msg == "") "AUTO" else "AMBIGUOUS", message = msg)
}
