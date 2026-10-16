#' Bank parser registry
#'
#' Every file in Banks/ calls register_bank() once per statement format it
#' supports. The engine (and the Excel bank drop-down, via bank_registry.csv)
#' only knows banks through this registry, so adding a bank = adding a file.

.bsp_registry <- new.env(parent = emptyenv())

#' @param id          short code, also used as the output sheet name (e.g. "BRI")
#' @param name        display name shown in Excel
#' @param variant     statement format name (a bank may register several)
#' @param detect      function(doc) -> score between 0 and 1
#' @param parse       function(doc) -> new_statement(...)
#' @param description free text for users (which statement this variant reads)
register_bank <- function(id, name, detect, parse, variant = "Standard",
                          description = "") {
  key <- paste(id, variant, sep = "|")
  assign(key, list(id = id, name = name, variant = variant, detect = detect,
                   parse = parse, description = description),
         envir = .bsp_registry)
  invisible(key)
}

registered_banks <- function() {
  keys <- sort(ls(.bsp_registry))
  lapply(keys, function(k) get(k, envir = .bsp_registry))
}

#' Registry as a table (written to bank_registry.csv for the Excel UI)
bank_registry_table <- function() {
  b <- registered_banks()
  data.frame(bank_id = vapply(b, `[[`, "", "id"),
             bank_name = vapply(b, `[[`, "", "name"),
             variant = vapply(b, `[[`, "", "variant"),
             description = vapply(b, `[[`, "", "description"),
             stringsAsFactors = FALSE)
}

#' Look up a parser; variant may be "" / NA to take the bank's only/first variant
get_bank <- function(id, variant = NA) {
  b <- Filter(function(x) toupper(x$id) == toupper(id), registered_banks())
  if (length(b) == 0) stop_bsp("BANK_UNKNOWN", paste0("Parser bank '", id, "' tidak tersedia"))
  if (!is.na(variant) && variant != "" && toupper(variant) != "AUTO") {
    v <- Filter(function(x) toupper(x$variant) == toupper(variant), b)
    if (length(v) == 0) stop_bsp("VARIANT_UNKNOWN",
                                 paste0("Format '", variant, "' untuk bank ", id, " tidak tersedia"))
    return(v[[1]])
  }
  b
}
